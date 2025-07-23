# MCP Client Configuration

## For VS Code/Cursor MCP Extension

Use this configuration in your MCP settings:

### Option 1: SSE Transport (Recommended)
```json
{
    "servers": {
        "excel-mcp-server": {
            "url": "https://excel-mcp-o7k4.onrender.com/sse",
            "type": "sse"
        }
    },
    "inputs": []
}
```

### Option 2: HTTP Transport
```json
{
    "servers": {
        "excel-mcp-server": {
            "url": "https://excel-mcp-o7k4.onrender.com",
            "type": "http"
        }
    },
    "inputs": []
}
```

### Option 3: If your client supports WebSocket
```json
{
    "servers": {
        "excel-mcp-server": {
            "url": "wss://excel-mcp-o7k4.onrender.com/ws",
            "type": "websocket"
        }
    },
    "inputs": []
}
```

## Troubleshooting

1. **404 Error**: Try different endpoints:
   - `/sse` for Server-Sent Events
   - `/` for HTTP
   - `/ws` for WebSocket

2. **Connection Issues**: 
   - Ensure the server is deployed and running
   - Check if the URL is accessible in browser
   - Try different transport types

3. **Authentication**: Currently no auth is required, but this can be added later.

## Testing the Server

You can test if the server is running by visiting:
- Main endpoint: https://excel-mcp-o7k4.onrender.com
- Health check: Use the health_check tool via MCP client
- SSE endpoint: https://excel-mcp-o7k4.onrender.com/sse

## Local Testing Configuration

For local testing (when running on localhost:8000):

```json
{
    "servers": {
        "excel-mcp-server-local": {
            "url": "http://localhost:8000/sse",
            "type": "sse"
        }
    },
    "inputs": []
}
```
