# google-slides-mcp

Tool for programmatically writing to Google Slides presentations using the [Model Context Protocol](https://modelcontextprotocol.io/)

## Run

```bash

node server.js write-slide <slide_num> <content_type> <text>
```

### Examples:

Writing a title to slide # 1

```bash
node server.js write-slide 1 title "This is the new title on slide 1"
```

Writing content to slide # 4

```bash
node server.js write-slide 4 content "This is the new content in the body of slide 4"
```

## Setup

1. Enable the Google Slides API in the Google Cloud console.

2. Download your client credentials to a file called `gcp-oauth.keys.json` in this directory.

3. In the `OAuth consent screen`, add your email to the Test users.

4. Install the required node.js packages listed in `package.json`

5. Now run the `write-slide` command in the format shown above. The first time you run, your browser should open for OAuth authentication with your email, and will then save your credentials to `.slides-server-credentials.json` locally so you won't have to re-authenticate again.
