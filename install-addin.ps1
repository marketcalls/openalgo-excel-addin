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
