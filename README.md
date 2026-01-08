# Excel to Kusto Datatable Converter

An Excel add-in that converts selected cells to Kusto datatable syntax.

This add-in is not published to the Office Add-ins Store, so it won't appear when you browse for
add-ins in Excel. Instead, you'll need to "sideload" it manually using a trusted catalog. This is a
one-time setup process that takes just a few minutes. You do not need to run a local server because
the files are hosted at
[microsoft.github.io/excel-to-kusto-datatable](https://microsoft.github.io/excel-to-kusto-datatable/).

Once sideloaded, the add-in works exactly like any other Excel add-in. It appears in the ribbon,
persists across sessions, and can be used whenever you need it.

## Quick Start

### Sideloading the Add-in

These instructions are for Windows, but if you have trouble or need other platforms (including
Office on the web), see the official Microsoft documentation on
[Sideloading Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

1. Download the [manifest.xml](https://microsoft.github.io/excel-to-kusto-datatable/manifest.xml) to
   your machine
2. Create a network share folder and place the manifest file in it
3. In Excel: **File → Options → Trust Center → Trust Center Settings**
4. Go to **Trusted Add-in Catalogs** → Add the network share path (e.g., `\\localhost\ExcelAddins`)
5. Restart Excel
6. Go to **Home → Add-ins → Advanced** → Select the add-in
7. Verify that the add-in button appears in the **Data** tab under the **Kusto** group

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
