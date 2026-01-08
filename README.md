# Excel to Kusto Datatable Converter

An Excel add-in that converts selected cells to Kusto datatable syntax.

## Quick Start

### Sideloading the Add-in

1. Save `manifest.xml` to a local folder (e.g., `C:\ExcelAddins\`)
2. Share that folder on your network
3. In Excel: **File → Options → Trust Center → Trust Center Settings**
4. Go to **Trusted Add-in Catalogs** → Add your network share path
5. Restart Excel
6. **Insert → My Add-ins → Shared Folder** → Select the add-in

### Usage

1. Select cells in Excel
2. Click the **"To Datatable"** button in the **Data** tab
3. The task pane opens with the Kusto datatable output
4. Copy the output to clipboard and paste into a Kusto query

## Contributing

See [Contributing](CONTRIBUTING.md) for more information.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of
Microsoft trademarks or logos is subject to and must follow
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion
or imply Microsoft sponsorship. Any use of third-party trademarks or logos are subject to those
third-party's policies.
