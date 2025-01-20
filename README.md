# OpenAlgo Excel Add-in

An Excel add-in for interacting with OpenAlgo trading API. This add-in provides custom functions to fetch and display trading account information directly in Excel.

## Features

- `OPENALGO.FUNDS()` - Fetches and displays fund details including:
  - Available Cash
  - Collateral
  - M2M Realized
  - M2M Unrealized
  - Utilised Debits

## Prerequisites

- Node.js (version 12 or higher)
- npm (comes with Node.js)
- Microsoft Excel (2016 or later)

## Installing the Add-in

### Windows Installation

1. **Prepare the Manifest**:
   - Download the `manifest.production.xml` file
   - Rename it to `manifest.xml` (optional)

2. **Method 1: Centralized Deployment (Recommended for Organizations)**:
   - Open Excel
   - Go to `File` > `Options` > `Trust Center` > `Trust Center Settings`
   - Select `Trusted Add-in Catalogs`
   - Add the catalog URL: `https://excel.openalgo.in`
   - Click OK and restart Excel
   - Go to `Insert` > `My Add-ins` > `Shared Folder`
   - Select "OpenAlgo Excel Add-in"

3. **Method 2: Manual Sideloading**:
   ```powershell
   # Create the add-in directory if it doesn't exist
   $addInFolder = "$env:USERPROFILE\AppData\Local\Microsoft\Office\16.0\Wef"
   New-Item -ItemType Directory -Force -Path $addInFolder

   # Copy the manifest
   Copy-Item manifest.production.xml "$addInFolder\manifest.xml"
   ```
   Then:
   - Open Excel
   - Go to `Insert` > `My Add-ins` > `OPENALGO`
   - Click to load the add-in

4. **Method 3: Network Share**:
   - Place the manifest file on a network share
   - Open Excel
   - Go to `File` > `Options` > `Trust Center` > `Trust Center Settings`
   - Click `Trusted Add-in Catalogs`
   - Add the network share path
   - Click OK and restart Excel
   - The add-in will appear in `Insert` > `My Add-ins`

### Troubleshooting Windows Installation

1. **Add-in Not Appearing**:
   - Verify manifest location:
     ```powershell
     dir "$env:USERPROFILE\AppData\Local\Microsoft\Office\16.0\Wef\manifest.xml"
     ```
   - Check Excel version (must be 2016 or later)
   - Clear Office cache:
     ```powershell
     Remove-Item "$env:USERPROFILE\AppData\Local\Microsoft\Office\16.0\Wef\*" -Force -Recurse
     Remove-Item "$env:USERPROFILE\AppData\Local\Microsoft\Office\16.0\WebExtCache\*" -Force -Recurse
     ```

2. **Security Settings**:
   - Open Excel
   - Go to `File` > `Options` > `Trust Center` > `Trust Center Settings`
   - Enable:
     - `Allow add-ins to start`
     - `Allow web add-ins to work on trusted documents`
     - `Turn on Preview Features`

3. **Network Issues**:
   - Ensure https://excel.openalgo.in is accessible
   - Check your firewall settings
   - Try using a different network connection

4. **Reset if Nothing Works**:
   ```powershell
   # Full reset of Office Add-in cache
   Stop-Process -Name "Excel" -Force -ErrorAction SilentlyContinue
   Remove-Item "$env:USERPROFILE\AppData\Local\Microsoft\Office\16.0\Wef\*" -Force -Recurse
   Remove-Item "$env:USERPROFILE\AppData\Local\Microsoft\Office\16.0\WebExtCache\*" -Force -Recurse
   ```
   Then:
   - Restart Excel
   - Try installing the add-in again

### Quick Installation Script (Windows)

Save this as `install-addin.ps1`:
```powershell
# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning "Please run as administrator"
    exit
}

# Create directories
$wefPath = "$env:USERPROFILE\AppData\Local\Microsoft\Office\16.0\Wef"
New-Item -ItemType Directory -Force -Path $wefPath

# Copy manifest
Copy-Item manifest.production.xml "$wefPath\manifest.xml" -Force

# Clear cache
Remove-Item "$env:USERPROFILE\AppData\Local\Microsoft\Office\16.0\WebExtCache\*" -Force -Recurse -ErrorAction SilentlyContinue

Write-Host "Add-in installed successfully. Please restart Excel."
```

Run the script:
```powershell
powershell -ExecutionPolicy Bypass -File install-addin.ps1
```

### Mac OS Installation

1. **Create Add-in Directory**:
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
   ```

2. **Copy Manifest**:
   - Copy the `manifest.xml` file to the wef directory
   - Make sure to use the production manifest that points to excel.openalgo.in
   ```bash
   cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
   ```

3. **Open Excel**:
   - Launch Microsoft Excel
   - Create a new workbook or open an existing one
   - If Excel was already open, close and reopen it

4. **Enable Add-in**:
   - Go to Excel > Preferences
   - Click 'Security & Privacy'
   - Check 'Trust access to the Office Add-ins platform'
   - Click 'OK'

5. **Access the Add-in**:
   - Go to the 'Insert' tab
   - Click 'My Add-ins' (Office Add-ins)
   - Look for "OpenAlgo" in the list
   - Click to load the add-in

6. **Using the Add-in**:
   - The taskpane will open on the right
   - Enter your API configuration
   - Use the `OPENALGO.FUNDS()` function in any cell

### Troubleshooting Mac Installation

If the add-in doesn't appear:
1. Verify the manifest.xml is in the correct location:
   ```bash
   ls ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml
   ```

2. Check manifest.xml content:
   - Ensure all URLs point to https://excel.openalgo.in/
   - Verify there are no localhost URLs

3. Excel Security:
   - Excel > Preferences > Security & Privacy
   - Ensure all required permissions are granted

4. Clear Office Cache:
   ```bash
   rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/*
   ```

5. Restart Excel:
   - Completely close Excel
   - Reopen Excel
   - Try loading the add-in again

## Configuration

1. Open the add-in taskpane in Excel
2. Enter your OpenAlgo API credentials:
   - API Host (e.g., http://127.0.0.1:5000)
   - API Key

## Usage

After configuration, you can use the following functions in Excel:

```excel
=OPENALGO.FUNDS()
```

This will return a table with your account's fund details.

## Deployment Options

### 1. Development Mode
During development, the add-in runs on a local server:
```bash
npm start
```
This starts a development server at https://localhost:3000

### 2. Production Deployment
For production use, you have several options:

a) **Static Web Hosting**:
- Build the project: `npm run build`
- Deploy the `dist` folder to any static web hosting service:
  - Azure Static Web Apps
  - GitHub Pages
  - Netlify
  - Vercel
  - AWS S3 + CloudFront
- Update the manifest.xml with your production URLs
- Share the updated manifest with users

b) **SharePoint Deployment**:
- Deploy to SharePoint for organizational use
- Use SharePoint's built-in HTTPS
- Centralized distribution via Microsoft 365 admin center

c) **Office Add-in Store**:
- Package the add-in
- Submit to the Office Store
- Users can install directly from Excel's Add-in store

### 3. Centralized Deployment
For enterprise deployment:
1. Build the production version: `npm run build`
2. Deploy to your hosting solution
3. Update manifest.xml with production URLs
4. Use Microsoft 365 admin center to deploy to your organization
5. Users will automatically get the add-in in Excel

### Office Store Publishing
To publish your add-in to the Office Store:

1. **Prepare Your Hosting**:
   - Build your add-in: `npm run build`
   - Deploy the built files to a web hosting service with HTTPS
   - Common options:
     - Azure Static Web Apps (recommended)
     - Azure Web Apps
     - AWS S3 + CloudFront
     - Any web host with HTTPS support

2. **Update Manifest**:
   - Update manifest.xml with your production URLs
   - Ensure all URLs use HTTPS
   - Example:
     ```xml
     <SourceLocation DefaultValue="https://your-domain.com/taskpane.html"/>
     ```

3. **Submit to Office Store**:
   - Go to [Partner Center](https://partner.microsoft.com/dashboard/office/overview)
   - Create a new submission
   - Upload your manifest
   - Provide store listing details (descriptions, screenshots, etc.)
   - Submit for validation

4. **After Approval**:
   - Your add-in appears in Excel's Add-in store
   - Users can install with a few clicks
   - No need for manual manifest deployment
   - Updates are handled through the store

The server hosting your add-in files needs to be maintained, but users don't need to run any local servers - they just install from Excel's Add-in store and your hosted files are served from your production server.

For more details on deployment options, see [Office Add-ins deployment guide](https://docs.microsoft.com/en-us/office/dev/add-ins/publish/publish).

### Vercel Deployment

1. **Prerequisites**:
   - Install Vercel CLI: `npm i -g vercel`
   - Have a Vercel account
   - Own the domain (excel.openalgo.in)

2. **Deploy to Vercel**:
   ```bash
   # Login to Vercel if not already logged in
   vercel login

   # Deploy to Vercel (this will automatically run the build)
   vercel

   # After testing in preview URL, deploy to production
   vercel --prod
   ```

3. **What Happens During Deployment**:
   - Vercel detects the project as a Node.js project
   - Runs `npm run build` automatically (configured in vercel.json)
   - Builds the project and creates the dist directory
   - Serves files from the dist directory

4. **Configure Custom Domain**:
   - Go to your Vercel project settings
   - Navigate to "Domains"
   - Add domain: `excel.openalgo.in`
   - Follow Vercel's instructions to configure DNS settings
   - Wait for DNS propagation

5. **Verify Deployment**:
   - Visit https://excel.openalgo.in/taskpane.html
   - Ensure all assets load correctly
   - Test the add-in using the updated manifest.xml

6. **Automatic Deployments**:
   - Connect your GitHub repository to Vercel
   - Each push to main will trigger:
     1. A fresh build of the project
     2. Deployment of the built files
   - Preview deployments for pull requests

Note: The manifest.xml has been updated to use https://excel.openalgo.in/. Make sure your domain is properly configured before submitting to the Office Store.

### Production Files
When deploying to production, you need to host the following files from your `dist` folder:

1. **Required Files**:
   ```
   dist/
   ├── taskpane.html       # Main UI
   ├── taskpane.js         # UI functionality
   ├── functions.js        # Excel functions
   ├── functions.json      # Function metadata
   ├── polyfill.js        # Browser compatibility
   ├── *.css              # Styles
   └── assets/            # Icons
       ├── icon-16.png
       ├── icon-32.png
       ├── icon-64.png
       ├── icon-80.png
       └── icon-128.png
   ```

2. **Optional Files** (for debugging):
   - Source maps (*.map files)
   - License files (*.LICENSE.txt)

3. **Do Not Host**:
   - `manifest.xml` - This is used for add-in registration, not hosted
   - Source files from `src/` directory
   - Development files (package.json, webpack.config.js, etc.)

To deploy:
1. Run `npm run build` to generate the `dist` folder
2. Upload all required files to your web host
3. Maintain the same file structure as in the `dist` folder
4. Ensure all files are served with correct MIME types
5. Enable HTTPS for all files

## Development

- `npm start` - Starts the development server
- `npm run build` - Builds the production version
- `npm run lint` - Runs linting
- `npm test` - Runs tests

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Built with [Office Add-ins Framework](https://docs.microsoft.com/en-us/office/dev/add-ins/)
