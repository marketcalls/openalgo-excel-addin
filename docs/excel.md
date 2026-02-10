# Excel

## OpenAlgo Excel Add-in Documentation

### Introduction

OpenAlgo provides seamless integration with Microsoft Excel through an Office Web Add-in for fetching trading account information directly in your spreadsheet. The add-in uses custom Excel functions under the `OPENALGO` namespace and a taskpane for API configuration.

This is a cross-platform add-in that works on **Windows**, **Mac**, and **Excel for the Web** (Excel 2016 or later required).

---

## Install the OpenAlgo Excel Add-in

The OpenAlgo Excel Add-in is a web-based Office Add-in hosted at `https://excel.openalgo.in`. Unlike traditional Excel add-ins, there is **no separate 32-bit or 64-bit version** -- a single add-in works across all platforms.

### Prerequisites

- Microsoft Excel 2016 or later (Desktop or Web)

### Windows Installation

#### Method 1: Manual Sideloading (Recommended)

1. **Download the manifest file**:
   - Download `manifest.production.xml` from the [GitHub releases page](https://github.com/marketcalls/openalgo-excel-addin/releases)

2. **Copy to the Office Add-in folder**:
   ```powershell
   # Create the add-in directory if it doesn't exist
   $addInFolder = "$env:USERPROFILE\AppData\Local\Microsoft\Office\16.0\Wef"
   New-Item -ItemType Directory -Force -Path $addInFolder

   # Copy the manifest
   Copy-Item manifest.production.xml "$addInFolder\manifest.xml"
   ```

3. **Load the add-in in Excel**:
   - Open Excel
   - Go to `Insert` > `My Add-ins`
   - Select **OpenAlgo Excel Add-in**
   - Click to load

#### Method 2: Trusted Catalog

1. Open Excel
2. Go to `File` > `Options` > `Trust Center` > `Trust Center Settings`
3. Select `Trusted Add-in Catalogs`
4. Add the catalog URL: `https://excel.openalgo.in`
5. Click OK and restart Excel
6. Go to `Insert` > `My Add-ins` > `Shared Folder`
7. Select **OpenAlgo Excel Add-in**

#### Method 3: Network Share

1. Place the `manifest.production.xml` file on a network share
2. Open Excel > `File` > `Options` > `Trust Center` > `Trust Center Settings`
3. Click `Trusted Add-in Catalogs`
4. Add the network share path
5. Click OK and restart Excel
6. The add-in will appear in `Insert` > `My Add-ins`

#### Quick Installation Script (Windows)

Save and run the following as `install-addin.ps1`:

```powershell
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

#### Troubleshooting Windows Installation

1. **Add-in Not Appearing**:
   - Verify the manifest exists:
     ```powershell
     dir "$env:USERPROFILE\AppData\Local\Microsoft\Office\16.0\Wef\manifest.xml"
     ```
   - Ensure your Excel version is 2016 or later
   - Clear the Office cache:
     ```powershell
     Remove-Item "$env:USERPROFILE\AppData\Local\Microsoft\Office\16.0\Wef\*" -Force -Recurse
     Remove-Item "$env:USERPROFILE\AppData\Local\Microsoft\Office\16.0\WebExtCache\*" -Force -Recurse
     ```

2. **Security Settings**:
   - Go to `File` > `Options` > `Trust Center` > `Trust Center Settings`
   - Enable:
     - `Allow add-ins to start`
     - `Allow web add-ins to work on trusted documents`

3. **Network Issues**:
   - Ensure `https://excel.openalgo.in` is accessible from your network
   - Check your firewall settings

4. **Full Reset**:
   ```powershell
   Stop-Process -Name "Excel" -Force -ErrorAction SilentlyContinue
   Remove-Item "$env:USERPROFILE\AppData\Local\Microsoft\Office\16.0\Wef\*" -Force -Recurse
   Remove-Item "$env:USERPROFILE\AppData\Local\Microsoft\Office\16.0\WebExtCache\*" -Force -Recurse
   ```
   Then restart Excel and install the add-in again.

---

### Mac Installation

1. **Create the add-in directory**:
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
   ```

2. **Copy the manifest**:
   ```bash
   cp manifest.production.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml
   ```

3. **Enable add-ins in Excel**:
   - Open Excel
   - Go to `Excel` > `Preferences` > `Security & Privacy`
   - Check `Trust access to the Office Add-ins platform`
   - Click OK

4. **Load the add-in**:
   - Go to the `Insert` tab
   - Click `My Add-ins` (Office Add-ins)
   - Select **OpenAlgo** from the list

#### Troubleshooting Mac Installation

- Verify the manifest is in the correct location:
  ```bash
  ls ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml
  ```
- Clear the Office cache:
  ```bash
  rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/*
  ```
- Restart Excel completely and try again

---

### Configuration

After installing the add-in, you need to configure your OpenAlgo API connection:

1. Open the add-in taskpane by clicking **Show OpenAlgo** on the Home tab
2. Enter your **API Host** (default: `http://127.0.0.1:5000`)
3. Enter your **API Key**
4. Click **Save Settings**

The settings are saved per workbook and will persist across sessions.

---

### Account Functions

#### Retrieve Funds

**Function:** `OPENALGO.FUNDS()`

**Example Usage:**

```excel
=OPENALGO.FUNDS()
```

**Returns:** A 2D array (table) with fund and margin details from your trading account:

| Field            | Description                        |
|------------------|------------------------------------|
| Available Cash   | Cash available for trading         |
| Collateral       | Collateral value in account        |
| M2M Realized     | Mark-to-market realized profit/loss |
| M2M Unrealized   | Mark-to-market unrealized profit/loss |
| Utilised Debits  | Total debits utilized              |

**Error Handling:** If the API call fails, the function returns `[['Error', '<error message>']]` in the cell range.

---

### Notes

- The add-in requires an active internet connection to load from `https://excel.openalgo.in`
- Your OpenAlgo server must be accessible from the machine running Excel
- API settings are stored per workbook -- each workbook can have its own configuration
- Test in **OpenAlgo Analyzer Mode** before using in live markets

### Support

For more details, visit [OpenAlgo Docs](https://docs.openalgo.in/).
