import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  ListToolsRequestSchema,
  CallToolRequestSchema
} from "@modelcontextprotocol/sdk/types.js";
import { google } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';
import http from 'http';
import open from 'open';
import { promisify } from 'util';
import { z } from "zod";
import fs from 'fs/promises';

const PRESENTATION_ID = 'TODO-PUT-GOOGLE-SLIDES-PRESENTATION-ID-HERE';
const CREDENTIALS_PATH = 'gcp-oauth.keys.json';
const TOKEN_PATH = '.slides-server-credentials.json';
const SCOPES = ['https://www.googleapis.com/auth/presentations'];

const PORT = 4892;

const WriteTitleSchema = z.object({
  text: z.string()
});


const server = new Server({
  name: "slides-writer",
  version: "1.0.0"
}, {
  capabilities: {
    tools: {}
  }
});

async function authenticate(credentials) {
  const { client_secret, client_id } = credentials.installed;
  const oauth2Client = new OAuth2Client(
    client_id,
    client_secret,
    `http://localhost:${PORT}`
  );

  return new Promise((resolve, reject) => {
    console.log('Setting up callback server...');
    const server = http
      .createServer(async (req, res) => {
        console.log('Received callback request:', req.url);
        try {
          if (req.url?.includes('code=')) {
            const qs = new URL(req.url, `http://localhost:${PORT}`).searchParams;
            const code = qs.get('code');
            console.log('Got auth code, getting tokens...');

            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end('<h1>Authentication successful! You can close this window.</h1>');

            const { tokens } = await oauth2Client.getToken(code);
            console.log('Got tokens, saving...');
            
            oauth2Client.credentials = tokens;
            await fs.writeFile(TOKEN_PATH, JSON.stringify(tokens));
            
            // Force close all connections
            server.closeAllConnections();
            await promisify(server.close.bind(server))();
            console.log('Callback server closed');
            resolve(oauth2Client);
          }
        } catch (e) {
          console.error('Error handling callback:', e);
          server.closeAllConnections();
          server.close();
          reject(e);
        }
      })
      .listen(PORT, () => {
        console.log(`Listening for OAuth callback on port ${PORT}`);
        const authorizeUrl = oauth2Client.generateAuthUrl({
          access_type: 'offline',
          scope: SCOPES,
        });
        console.log('Opening auth URL:', authorizeUrl);
        open(authorizeUrl, { wait: false });
      });

    server.on('error', (e) => {
      console.error('Server error:', e);
      server.closeAllConnections();
      server.close();
      reject(e);
    });
  });
}

async function getSlideService() {
  try {
    // Load client credentials
    console.log('Loading client credentials...');
    const content = await fs.readFile(CREDENTIALS_PATH);
    const credentials = JSON.parse(content.toString());

    // Try to load existing token
    try {
      console.log('Checking for existing token...');
      const token = await fs.readFile(TOKEN_PATH);
      const { client_secret, client_id } = credentials.installed;
      const oauth2Client = new OAuth2Client(client_id, client_secret, `http://localhost:${PORT}`);
      oauth2Client.credentials = JSON.parse(token.toString());
      return google.slides({ version: 'v1', auth: oauth2Client });
    } catch (e) {
      console.log('No existing token, starting auth flow...');
      const oauth2Client = await authenticate(credentials);
      return google.slides({ version: 'v1', auth: oauth2Client });
    }
  } catch (error) {
    console.error('Error in auth flow:', error);
    throw error;
  }
}

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  if (name === "write-slide-title") {
    try {
      const { text } = WriteTitleSchema.parse(args);
      console.log("Writing text:", text);
      
      const slides = await getSlideService();
      const titleId = await findTitleShape(slides, PRESENTATION_ID);
      
      console.log("Found title element with ID:", titleId);

      // Update the text
      const result = await slides.presentations.batchUpdate({
        presentationId: PRESENTATION_ID,
        requestBody: {
          requests: [
            // First delete existing text
            {
              deleteText: {
                objectId: titleId,
                textRange: {
                  type: 'ALL'
                }
              }
            },
            // Then insert new text
            {
              insertText: {
                objectId: titleId,
                text: text,
                insertionIndex: 0
              }
            }
          ]
        }
      });

      console.log("Update result:", result.data);

      return {
        content: [{
          type: "text",
          text: `Successfully wrote "${text}" to slide title`
        }]
      };

    } catch (error) {
      console.error("Error:", error);
      return {
        isError: true,
        content: [{
          type: "text",
          text: `Error: ${error.message}`
        }]
      };
    }
  }

  throw new Error(`Unknown tool: ${name}`);
});

async function findSlideElements(slides, presentationId, slideIndex) {
  const presentation = await slides.presentations.get({
    presentationId: PRESENTATION_ID,
  });
  
  const slide = presentation.data.slides[slideIndex - 1];
  if (!slide || !slide.pageElements) {
    throw new Error(`Could not find slide number ${slideIndex}`);
  }
  
  console.log(`Found slide elements for slide ${slideIndex}:`, slide.pageElements);
  
  // Find the elements by their vertical position
  // The element with higher translateY is the content (appears lower on slide)
  const sortedElements = slide.pageElements
    .filter(el => el.shape?.shapeType === 'TEXT_BOX')
    .sort((a, b) => a.transform.translateY - b.transform.translateY);

  if (sortedElements.length < 2) {
    throw new Error("Could not find required text elements on slide");
  }

  return {
    titleId: sortedElements[0].objectId,  // Higher up element is title
    contentId: sortedElements[1].objectId  // Lower element is content
  };
}

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [{
      name: "write-slide-title",
      description: "Write the category and amount text at the top of the slide",
      inputSchema: {
        type: "object",
        properties: {
          text: {
            type: "string",
            description: "The title text to write (e.g. 'CATEGORY $400')"
          }
        },
        required: ["text"]
      }
    },
    {
      name: "write-slide-content",
      description: "Write the clue text in the center of the slide",
      inputSchema: {
        type: "object",
        properties: {
          text: {
            type: "string",
            description: "The clue text to write"
          }
        },
        required: ["text"]
      }
    }]
  };
});

async function replaceTextInShape(slides, elementId, newText) {
  try {
    // First try to get the current text to see if it's empty
    const presentation = await slides.presentations.get({
      presentationId: PRESENTATION_ID,
      fields: 'slides'
    });
    
    // Try to delete existing text only if there is text
    const hasText = presentation.data.slides.some(slide =>
      slide.pageElements?.some(el => 
        el.objectId === elementId && 
        el.shape?.text?.textElements?.length > 0
      )
    );

    if (hasText) {
      // Delete existing text
      await slides.presentations.batchUpdate({
        presentationId: PRESENTATION_ID,
        requestBody: {
          requests: [{
            deleteText: {
              objectId: elementId,
              textRange: {
                type: 'ALL'
              }
            }
          }]
        }
      });
    }

    // Insert new text
    await slides.presentations.batchUpdate({
      presentationId: PRESENTATION_ID,
      requestBody: {
        requests: [{
          insertText: {
            objectId: elementId,
            text: newText,
            insertionIndex: 0
          }
        }]
      }
    });

  } catch (error) {
    console.error('Error in replaceTextInShape:', error);
    throw error;
  }
}

// Add command line handling
if (process.argv[2] === 'auth') {
  console.log("Starting authentication flow...");
  getSlideService().then(() => {
    console.log("Authentication completed successfully!");
    process.exit(0);
  }).catch(error => {
    console.error("Authentication failed:", error);
    process.exit(1);
  });
} 
// Update command line handling to use slide numbers
if (process.argv[2] === 'write-slide') {
  if (process.argv.length < 6) {
    console.log("Usage: node server.js write-slide <slide-number> title|content <text>");
    process.exit(1);
  }

  const slideNumber = parseInt(process.argv[3], 10);
  if (isNaN(slideNumber) || slideNumber < 1) {
    console.error("Slide number must be a positive integer");
    process.exit(1);
  }

  const textType = process.argv[4];
  const newText = process.argv[5];

  console.log(`Writing ${textType} text to slide ${slideNumber}...`);
  getSlideService().then(async (slides) => {
    try {
      const { titleId, contentId } = await findSlideElements(slides, PRESENTATION_ID, slideNumber);
      const elementId = textType === 'title' ? titleId : contentId;
      
      await replaceTextInShape(slides, elementId, newText);
      console.log("Update successful!");
      process.exit(0);
    } catch (error) {
      console.error("Error updating slide:", error);
      process.exit(1);
    }
  });
} else {
  // Start MCP server
  const transport = new StdioServerTransport();
  server.connect(transport).catch(error => {
    console.error("Fatal error:", error);
    process.exit(1);
  });
}