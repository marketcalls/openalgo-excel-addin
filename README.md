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

## Installation

1. Clone the repository:
```bash
git clone https://github.com/marketcalls/openalgo-excel-addin.git
cd openalgo-excel-addin
```

2. Install dependencies:
```bash
npm install
```

3. Start the development server:
```bash
npm start
```

4. Sideload the add-in in Excel:
   - Open Excel
   - Go to Insert > My Add-ins > Upload My Add-in
   - Select the manifest file from the project

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
