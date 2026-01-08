# Excel to Kusto Datatable Converter

## Project Overview

An Excel add-in that allows users to select cells and quickly convert them to Kusto datatable syntax. The tool prioritizes ease of use and clean, maintainable code suitable for open-source distribution.

---

## Goals

- **Simple UX**: One-click conversion from selected Excel cells to Kusto datatable format
- **Ribbon Access**: Available via custom ribbon button (context menu not supported in Office.js)
- **Flexible Output**: Display in task pane and copy directly to clipboard
- **Clean Codebase**: Readable, well-organized TypeScript code suitable for GitHub publication
- **Easy Distribution**: Hosted on GitHub Pages, users sideload a manifest file
- **Cross-Platform**: Works on Windows, Mac, and Excel Online

---

## Requirements

### Functional Requirements

| ID | Requirement | Priority |
|----|-------------|----------|
| F1 | User can select a range of cells in Excel | Must Have |
| F2 | Convert selection to Kusto datatable syntax | Must Have |
| F3 | Access conversion via ribbon button | Must Have |
| F4 | Display result in task pane | Must Have |
| F5 | Copy result to clipboard | Must Have |
| F6 | Handle empty cells gracefully | Must Have |
| F7 | Preserve data types where applicable | Must Have |

### Non-Functional Requirements

| ID | Requirement |
|----|-------------|
| NF1 | Written in TypeScript |
| NF2 | Compatible with Excel 2016+, Microsoft 365, and Excel Online |
| NF3 | Hosted on GitHub Pages (free, no server maintenance) |
| NF4 | Clean separation of concerns (but not over-engineered) |
| NF5 | Meaningful inline comments without excess verbosity |

---

## Technical Approach

### Add-in Technology: Office.js (Web Add-in)

**Why Office.js:**
- Cross-platform: Windows, Mac, and Excel Online
- Free hosting via GitHub Pages
- Auto-updates: push to repo → users get new version immediately
- Fork-friendly: others can fork and have their own working instance
- Modern web stack (TypeScript, HTML, CSS)

**Trade-offs:**
- No right-click context menu support (ribbon only)
- Task pane UI instead of popup dialog
- Async/Promise-based API
- Users must sideload manifest (one-time setup)

### Tech Stack

- **TypeScript** - Type safety, better tooling
- **Office.js API** - Excel interaction
- **GitHub Pages** - Free static hosting with HTTPS
- **Node.js** - Build tooling only (not runtime)

---

## Project Structure

```
excel-to-kusto-datatable/
├── manifest.xml              # Office Add-in manifest (points to GitHub Pages URL)
├── package.json              # Node.js dependencies and scripts
├── tsconfig.json             # TypeScript configuration
├── webpack.config.js         # Build configuration
├── README.md
├── PLANNING.md               # This document
├── LICENSE
│
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html     # Task pane UI structure
│   │   ├── taskpane.ts       # Task pane logic and event handlers
│   │   └── taskpane.css      # Styling
│   │
│   └── services/
│       └── datatableConverter.ts  # Core conversion logic (pure functions)
│
└── dist/                     # Built files (deployed to GitHub Pages)
    ├── taskpane.html
    ├── taskpane.js
    └── taskpane.css
```

### Design Notes

- **manifest.xml**: Defines the add-in, ribbon buttons, and GitHub Pages URL
- **taskpane.ts**: Handles UI events, reads Excel selection, delegates to converter
- **datatableConverter.ts**: Pure conversion logic, no Office.js dependencies, easily testable
- **dist/**: Production build output, served by GitHub Pages

This structure keeps things organized without over-abstracting. No frameworks like React/Angular - just vanilla TypeScript with clean separation where it makes sense.

---

## Kusto Datatable Syntax

*Document the expected output format here. Example:*

```
datatable(Column1:string, Column2:real, Column3:datetime)
[
    "Value1", 123, datetime(2024-01-15),
    "Value2", real(null), datetime(01/16/2024),
]
```

### Conversion Rules

1. First row treated as headers by default (configurable via checkbox)
2. Infer data types from cell values/formats:
   - Whole numbers → `long`
   - Decimal numbers → `real`
   - Dates (string format) → `datetime`
   - Everything else → `string`
3. **Schema Mapping**: User can override inferred types before conversion
   - All Kusto types supported: `string`, `long`, `real`, `datetime`, `bool`, `guid`, `timespan`, `decimal`, `dynamic`
   - Useful for Excel serial dates (numbers) that should be `datetime`
4. String values wrapped in double quotes
5. Proper escaping for special characters
6. Empty/null cells → typed nulls (e.g., `datetime(null)`, `long(null)`, `real(null)`, `""`)
7. Ambiguous dates → pass through to `datetime()` operator as-is (e.g., `datetime(01/07/2026)`)

---

## User Interface

### Ribbon Button

- **Tab**: Added to existing "Data" tab (less intrusive than a custom tab)
- **Group**: "Kusto" group within the Data tab
- **Button**: "To Datatable" with table/code icon
- **Behavior**: Opens task pane on the right side of Excel
- **Tooltip**: "Convert selected cells to Kusto datatable syntax"

### Task Pane (Right Sidebar)

```
┌─────────────────────────────────────┐
│  Kusto Datatable Converter          │
├─────────────────────────────────────┤
│                                     │
│  [✓] First row is header            │
│                                     │
│  [ Load Selection ]                 │
│                                     │
│  ┌─ Column Types ─────────────────┐ │
│  │ Name        [string     ▼]     │ │
│  │ StartDate   [datetime   ▼]     │ │
│  │ Amount      [real       ▼]     │ │
│  │ IsActive    [bool       ▼]     │ │
│  └────────────────────────────────┘ │
│                                     │
│  [ Convert to Datatable ]           │
│                                     │
│  ┌───────────────────────────────┐  │
│  │ datatable(Name:string,        │  │
│  │           StartDate:datetime, │  │
│  │           Amount:real,        │  │
│  │           IsActive:bool)      │  │
│  │ [                             │  │
│  │     "Alice", datetime(...),   │  │
│  │     50.5, true,               │  │
│  │ ]                             │  │
│  └───────────────────────────────┘  │
│                                     │
│  [ Copy to Clipboard ]              │
│                                     │
│  ✓ Copied to clipboard!             │
│                                     │
└─────────────────────────────────────┘
```

**Two-Step Conversion Flow:**
1. **Load Selection** - Reads Excel selection and infers column types
2. **Schema Mapping** - User reviews/adjusts column types via dropdowns
3. **Convert** - Generates datatable with user-specified types

**UI Elements:**
- Checkbox for "First row is header" (default: checked)
- "Load Selection" button reads current Excel selection and shows schema
- Column type dropdowns allow changing any column to any Kusto type
- "Convert to Datatable" generates the output
- Output area is read-only, scrollable, monospace font
- "Copy to Clipboard" button with status confirmation
- Task pane stays open for repeated use

---

## Build & Installation Instructions

### For Users (Installing the Add-in)

📺 **Sideloading instructions**: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins

**Quick Install Steps:**

1. Download `manifest.xml` to a local folder (e.g., `C:\ExcelAddins\`)
2. Right click on the folder and choose "Give access to > Specific people..." to share it on your network.
3. Excel → File → Options → Trust Center → Trust Center Settings
4. Trusted Add-in Catalogs → Add your network share path
5. Restart Excel
6. Home → Add-ins → Advanced → Shared Folder → Select the add-in

### For Developers (Building from Source)

**Prerequisites:**
- Node.js 18+ 
- npm or yarn
- Git

**Build Steps:**

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/excel-to-kusto-datatable.git
cd excel-to-kusto-datatable

# Install dependencies
npm install

# Build for production
npm run build

# Start local dev server (for testing)
npm run dev
```

**Local Development:**

1. Run `npm run dev` to start local server at `https://localhost:3000`
2. Sideload `manifest.xml` (points to localhost for dev)
3. Make changes → browser auto-refreshes

**Deploying to GitHub Pages:**

1. Push to `main` branch
2. GitHub Actions builds and deploys to GitHub Pages automatically
3. Users with manifest installed get updates on next Excel launch

### Uninstallation

1. In Excel: Insert → My Add-ins
2. Right-click the add-in → Remove

---

## Development Phases

### Phase 1: Hello World ✅
**Goal**: Minimal add-in running in Excel with local dev server

- [x] Initialize project (npm init, install dependencies)
- [x] Create manifest.xml pointing to localhost
- [x] Create minimal taskpane.html with "Hello World" text
- [x] Configure webpack and dev server with HTTPS
- [x] Add ribbon button to Data tab

**Verify**: 
1. ✅ Run `npm run dev` → dev server starts at https://localhost:3000
2. ✅ Sideload manifest into Excel
3. ✅ Click button in Data tab → task pane opens showing "Hello World"

---

### Phase 2: Read Selection ✅
**Goal**: Read selected cells and display raw data in task pane

- [x] Add "Convert" button to task pane
- [x] Implement Office.js code to read selected range
- [x] Display cell values as JSON or plain text in task pane

**Verify**:
1. Select some cells in Excel (e.g., A1:C3)
2. Click "Convert" in task pane
3. Task pane shows the raw cell values from your selection

---

### Phase 3: Core Conversion ✅
**Goal**: Convert selection to proper Kusto datatable syntax

- [x] Create datatableConverter.ts with conversion logic
- [x] Implement data type inference (string, long, real, datetime)
- [x] Generate proper datatable syntax with headers
- [x] Handle null/empty cells
- [x] Unit tests for converter (Jest)

**Verify**:
1. Select cells with headers and mixed data types
2. Click "Convert" → output shows valid Kusto datatable syntax
3. Run unit tests → all pass (`npm test`)

---

### Phase 4: Polish UI ✅
**Goal**: Complete task pane with all features

- [x] Add "First row is header" checkbox (default checked)
- [x] Style output area (monospace, scrollable)
- [x] Add "Copy to Clipboard" button with confirmation
- [x] Error handling with user-friendly messages
- [x] Schema mapping UI - show inferred column types with dropdown overrides
- [x] Support all Kusto data types (string, long, real, datetime, bool, guid, timespan, decimal, dynamic)
- [x] Two-step flow: Load Selection → Adjust Types → Convert

**Verify**:
1. Toggle checkbox → reloads and re-infers column types
2. Change column type dropdown → Convert uses new type
3. Click "Copy to Clipboard" → paste into text editor shows correct output
4. Select invalid range → shows helpful error message

---

### Phase 5: GitHub Deployment
**Goal**: Live add-in hosted on GitHub Pages

- [ ] Configure GitHub Pages on repository
- [ ] Set up GitHub Actions workflow for auto-build
- [ ] Create production manifest.xml with GitHub Pages URL
- [ ] Write README with installation instructions
- [ ] Test fresh install on different machine

**Verify**:
1. Push to main → GitHub Actions builds successfully
2. Visit https://yourusername.github.io/repo → files are served
3. Fresh Excel install with production manifest → add-in works
4. Test on Excel Online → add-in works

---


## Future Considerations (Out of Scope for v1)

- Settings persistence (remember preferences)
- Multiple output formats (other query languages)
- Keyboard shortcut binding
- Office Store publication (Microsoft AppSource)

---

## License

MIT License (recommended for open-source simplicity)

---

*Last Updated: January 7, 2026*
