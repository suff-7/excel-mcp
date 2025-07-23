# Excel FastMCP Server - Render Deployment

This guide explains how to deploy your Excel FastMCP server to Render.

## Prerequisites

1. A GitHub repository with your code
2. A Render account (free tier available)

## Deployment Steps

### Option 1: Using Render Dashboard

1. **Push your code to GitHub**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/your-username/excel-mcp-server.git
   git push -u origin main
   ```

2. **Connect to Render**
   - Go to [Render Dashboard](https://dashboard.render.com/)
   - Click "New +" → "Web Service"
   - Connect your GitHub repository
   - Select this repository

3. **Configure the service**
   - **Name**: `excel-mcp-server`
   - **Environment**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `python excel_fastmcp_server.py`
   - **Plan**: Free (or paid for production)

### Option 2: Using render.yaml (Infrastructure as Code)

1. Push your code with the `render.yaml` file
2. In Render Dashboard: "New +" → "Blueprint"
3. Connect your repository
4. Render will automatically configure based on `render.yaml`

## Environment Variables

The server automatically uses the `PORT` environment variable provided by Render.

## Health Check

The server includes a health check endpoint accessible via the `health_check` tool.

## File Upload Considerations

**Important**: Render's free tier has ephemeral storage. Uploaded Excel files will be lost when the service restarts. For production use:

1. Use Render's paid plans with persistent storage
2. Integrate with cloud storage (AWS S3, Google Cloud Storage, etc.)
3. Store files in a database

## Testing Your Deployment

Once deployed, your server will be available at:
```
https://your-app-name.onrender.com
```

You can test the health check and other endpoints using the MCP protocol.

## Troubleshooting

1. **Deployment fails**: Check the build logs in Render dashboard
2. **Server doesn't start**: Verify the start command and PORT configuration
3. **File operations fail**: Check file paths and permissions

## Production Considerations

1. Add authentication/authorization
2. Implement rate limiting
3. Add comprehensive error handling
4. Use environment variables for sensitive data
5. Set up monitoring and logging
6. Consider using a database for persistent storage

## Support

For issues specific to this Excel MCP server, check the logs in your Render dashboard.
