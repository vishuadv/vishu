# Northflank Deployment Guide

This guide will help you deploy the Loan Management System on Northflank using Neon PostgreSQL database.

## Prerequisites

1. **GitHub Repository**: Your code is already pushed to https://github.com/vishuadv/vishu
2. **Neon Database Account**: Create a free account at https://neon.tech
3. **Northflank Account**: Create a free account at https://northflank.com

## Step 1: Set up Neon Database

1. Sign up/login to [Neon Console](https://console.neon.tech)
2. Create a new project:
   - Click "Create a project"
   - Choose a name (e.g., "vishu")
   - Select a region closest to your users
   - Click "Create project"
3. Get your connection string:
   - Go to your project dashboard
   - Click "Connection Details"
   - Copy the connection string (format: `postgresql://username:password@ep-xxx.region.aws.neon.tech/neondb`)
4. Note: Replace `neondb` with `vishu` in the connection string

## Step 2: Deploy on Northflank

### Option A: Using GitHub Integration (Recommended)

1. Login to [Northflank](https://northflank.com)
2. Create a new account or login
3. Click "Create a service" → "Git service"
4. Connect your GitHub account
5. Select the `vishuadv/vishu` repository
6. Configure the service:
   - **Name**: vishu
   - **Branch**: main
   - **Build context**: /
   - **Dockerfile path**: Dockerfile
   - **Port**: 8080
7. Add environment variables:
   ```
   DATABASE_URL=postgresql://username:password@ep-xxx.region.aws.neon.tech/vishu
   SECRET_KEY=generate-a-secure-random-key
   ADMIN_USERNAME=shivam
   ADMIN_PASSWORD=Raaina@20
   FLASK_DEBUG=False
   PORT=8080
   ```
8. Set resources:
   - CPU: 0.5 vCPU
   - RAM: 512 MB
9. Click "Create service"

### Option B: Using Docker Image

1. Build and push Docker image to GitHub Container Registry:
   ```bash
   echo $GITHUB_TOKEN | docker login ghcr.io -u vishuadv --password-stdin
   docker build -t ghcr.io/vishuadv/vishu:latest .
   docker push ghcr.io/vishuadv/vishu:latest
   ```

2. On Northflank:
   - Create a new service → "Container service"
   - Image: `ghcr.io/vishuadv/vishu:latest`
   - Port: 8080
   - Add the same environment variables as above

## Step 3: Configure Environment Variables

In Northflank, add these environment variables to your service:

| Variable | Value | Description |
|----------|-------|-------------|
| `DATABASE_URL` | Your Neon connection string (postgresql://username:password@ep-xxx.region.aws.neon.tech/vishu) | PostgreSQL database URL |
| `SECRET_KEY` | Generate a secure key | Flask session encryption |
| `ADMIN_USERNAME` | shivam | Admin username |
| `ADMIN_PASSWORD` | Raaina@20 | Admin password (change in production) |
| `FLASK_DEBUG` | False | Disable debug mode |
| `PORT` | 8080 | Application port |

**Generate a secure SECRET_KEY:**
```python
import secrets
print(secrets.token_hex(32))
```

## Step 4: Initialize Database

The application will automatically create tables on first run. However, you may want to verify:

1. Check Northflank logs for any database connection errors
2. Access the application at the provided Northflank URL
3. Login with admin credentials

## Step 5: Access Your Application

After deployment, Northflank will provide:
- A public URL (e.g., `https://vishu-xxx.northflank.app`)
- Service logs
- Metrics and monitoring

## Troubleshooting

### Database Connection Issues
- Verify your Neon connection string is correct
- Check that Neon allows connections from Northflank's IP ranges
- Ensure the database name in the connection string matches

### Build Failures
- Check Northflank build logs
- Ensure Dockerfile is present and valid
- Verify all dependencies are in requirements.txt

### Application Errors
- Check service logs in Northflank
- Verify all environment variables are set
- Ensure PORT is set to 8080

## Security Recommendations

1. **Change default credentials**: Update ADMIN_USERNAME and ADMIN_PASSWORD
2. **Use strong SECRET_KEY**: Generate a random 64-character hex string
3. **Enable SSL**: Northflank provides HTTPS by default
4. **Regular backups**: Neon provides automatic backups
5. **Monitor logs**: Regularly check Northflank service logs

## Scaling

If you need to handle more traffic:

1. In Northflank service settings:
   - Increase CPU allocation (0.5 → 1 → 2 vCPU)
   - Increase RAM (512MB → 1GB → 2GB)
   - Add replicas (1 → 2 → 3 instances)

2. For Neon database:
   - Upgrade to paid tier for higher limits
   - Enable read replicas for better performance

## Cost Estimate

- **Northflank Free Tier**: 
  - 1 service with 0.5 vCPU, 512MB RAM
  - Sufficient for small to medium usage
  
- **Neon Free Tier**:
  - 0.5 GB storage
  - Sufficient for initial deployment

Both free tiers should handle moderate usage for the Loan Management System.
