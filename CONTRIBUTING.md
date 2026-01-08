# Contributing to Excel to Kusto Datatable Converter

This project welcomes contributions and suggestions. Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit
[https://cla.opensource.microsoft.com](https://cla.opensource.microsoft.com).

When you submit a pull request, a CLA bot will automatically determine whether you need to provide a
CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the
[Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more
information see the [Code of Conduct](CODE_OF_CONDUCT.md).

## Development Environment

You can contribute to this project from a Windows or macOS machine.

### Prerequisites

- Node.js 18+
- npm
- Git client and command line tools
- Excel (Desktop version - Windows or Mac)

### Quick Start

```sh
# Clone the repository
git clone https://github.com/Microsoft/excel-to-kusto-datatable.git
cd excel-to-kusto-datatable

# Install dependencies
npm install

# Start development server
npm run dev

# Run tests
npm test
```

### Local Development Setup

This project has two manifest files:

| File                 | Purpose                        | URLs Point To                                   |
| -------------------- | ------------------------------ | ----------------------------------------------- |
| `manifest.xml`       | Production (for users)         | `microsoft.github.io/excel-to-kusto-datatable/` |
| `manifest.local.xml` | Development (for contributors) | `localhost:3000`                                |

**To test locally:**

1. Run `npm run dev` to start the local dev server at `https://localhost:3000`
2. Sideload **`manifest.local.xml`** into Excel (not `manifest.xml`)
3. The add-in will appear as "Kusto Datatable (Local)" to distinguish it from the production version

**Sideloading instructions:**
https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing

> **Note:** Always use `manifest.local.xml` for development. The production `manifest.xml` points to
> GitHub Pages and won't reflect your local changes.

### Project Structure

```
src/
├── taskpane/
│   ├── taskpane.html         # Task pane markup
│   ├── taskpane.ts           # Task pane logic
│   └── taskpane.css          # Styles
├── services/
│   ├── datatableConverter.ts      # Conversion logic
│   └── datatableConverter.test.ts # Unit tests
└── assets/                   # Icons and images
```

## Pull Requests

### How to Create Pull Requests

Fork the project on GitHub and clone the upstream repo:

```sh
git clone https://github.com/Microsoft/excel-to-kusto-datatable.git
```

Navigate to the repo root:

```sh
cd excel-to-kusto-datatable
```

Add your fork as an origin:

```sh
git remote add fork https://github.com/YOUR_GITHUB_USERNAME/excel-to-kusto-datatable.git
```

Check out a new branch, make modifications, and push the branch to your fork:

```sh
git checkout -b feature
# edit files
git commit
git push fork feature
```

Open a pull request against the main `excel-to-kusto-datatable` repo.

### Tips and Best Practices

- If the PR is not ready for review, please mark it as
  [`draft`](https://github.blog/2019-02-14-introducing-draft-pull-requests/).
- Submit small, focused PRs addressing a single concern/issue.
- Make sure the PR title reflects the contribution.
- Write a summary that helps understand the change.
- Include screenshots or GIFs for UI changes.

### How to Get Pull Requests Merged

A PR is considered to be **ready to merge** when:

- All tests pass (`npm test`).
- Major feedback/comments are resolved.
- It has been open for review for at least one working day. This gives people reasonable time to
  review.
- Approved by at least one code owner.

### When to Start with a Discussion or Issue

**Submit a PR directly for:**

- Bug fixes with clear reproduction steps
- Obvious corrections (typos, broken links)
- Small improvements to existing functionality
- Adding or improving tests

**Start with a GitHub Discussion or Issue for:**

- New features or significant changes
- Changes to the conversion algorithm
- UI/UX redesigns
- Breaking changes to the output format

## Code Quality

### Running Tests

The project uses Jest for unit testing:

```sh
# Run all tests
npm test

# Run tests in watch mode during development
npm run test:watch

# Run tests with coverage report
npm run test:coverage
```

### Code Style

- Write TypeScript with strict type annotations
- Follow existing code patterns in the repository
- Use meaningful variable and function names
- Add JSDoc comments for public functions

### Testing Guidelines

- Write unit tests for new functionality in `*.test.ts` files
- Test edge cases (empty selections, special characters, different data types)
- Ensure existing tests pass before submitting a PR

## Building and Testing the Add-in

### Development Server

Start the development server to test changes in Excel:

```sh
npm run dev
```

The server runs at `https://localhost:3000`.

### SSL Certificate Setup (Windows)

Excel uses an embedded browser (WebView2) that doesn't share your browser's certificate exceptions.
You must install the dev server's self-signed certificate into your Windows trusted certificate
store.

#### Install the Certificate

After running `npm run dev` at least once (which generates the certificate), run these PowerShell
commands:

```powershell
# Navigate to project folder
cd path\to\excel-to-kusto-datatable

# Extract certificate from PEM file
$pem = Get-Content "node_modules\.cache\webpack-dev-server\server.pem" -Raw
$certMatch = [regex]::Match($pem, "-----BEGIN CERTIFICATE-----(\[\s\S\]*?)-----END CERTIFICATE-----")
$certPem = "-----BEGIN CERTIFICATE-----" + $certMatch.Groups[1].Value + "-----END CERTIFICATE-----"
$certPem | Out-File -FilePath "dev-server.crt" -Encoding ASCII

# Install to trusted root store
certutil -addstore -user -f "Root" "dev-server.crt"

# Set a friendly name (optional, helps identify it later)
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
$store.Open("ReadWrite")
$cert = $store.Certificates | Where-Object { $_.Subject -eq "CN=localhost" } | Select-Object -First 1
if ($cert) { $cert.FriendlyName = "Excel Add-in Dev Server (Kusto Datatable)" }
$store.Close()
```

#### Remove the Certificate Later

1. Press `Win + R`, type `certmgr.msc`, press Enter
2. Navigate to **Trusted Root Certification Authorities** → **Certificates**
3. Find "Excel Add-in Dev Server (Kusto Datatable)" (or "localhost")
4. Right-click → Delete

### Sideloading the Add-in into Excel

#### One-Time Setup

1. Copy `manifest.xml` to a local folder (e.g., `C:\ExcelAddins\`)
2. Right-click the folder → **Give access to** → **Specific people...** → Share it
3. In Excel: **File** → **Options** → **Trust Center** → **Trust Center Settings**
4. Go to **Trusted Add-in Catalogs**
5. Add your network share path (e.g., `\\YourPC\ExcelAddins`)
6. Check "Show in Menu" and click **OK**
7. Restart Excel

#### Load the Add-in

1. In Excel: **Home** → **Add-ins** → **More Add-ins**
2. Go to **Shared Folder** tab
3. Select **Kusto Datatable** and click **Add**

The add-in button appears in the **Data** tab under the **Kusto** group.

### Development Workflow

1. Start the dev server: `npm run dev`
2. Make changes to files in `src/`
3. The browser auto-refreshes; Excel task pane may need manual refresh (close and reopen)

### Production Build

Create a production build:

```sh
npm run build
```

Output files are generated in the `dist/` folder.

## Troubleshooting

### "Content blocked - not signed by valid security certificate"

The SSL certificate isn't trusted. Follow the
[SSL Certificate Setup](#ssl-certificate-setup-windows) steps above, then restart Excel.

### Add-in doesn't appear in Shared Folder

- Verify the network share is accessible (try `\\YourPC\ShareName` in File Explorer)
- Ensure `manifest.xml` is in the shared folder
- Restart Excel after adding the trusted catalog

### Changes not reflecting in Excel

- Close and reopen the task pane
- If ribbon buttons changed, remove and re-add the add-in
- Clear Office cache: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`

## Reporting Issues

When reporting issues, please include:

- Excel version and operating system
- Steps to reproduce the problem
- Expected behavior vs actual behavior
- Sample data (if applicable, anonymized)
- Screenshots or screen recordings for UI issues
