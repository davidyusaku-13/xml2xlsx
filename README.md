# xml2xlsx

Convert XML to XLSX. Auto-detects Microsoft Excel XML Spreadsheet format (Office 2003) and generic XML.

```
Usage: xml2xlsx <input.xml> [output.xlsx]
```

## Features

**Two parsers, auto-selected based on input format:**

### Excel XML Spreadsheet (`urn:schemas-microsoft-com:office:spreadsheet`)

Streaming token-based parser extracts each `<Worksheet>` as a sheet and each `<Row>` as a row. Uses `excelize.StreamWriter` for fast output.

- Multi-sheet workbooks supported
- Sheet names carried over from `ss:Name` attributes
- Cell values from `<Cell><Data>` elements
- Tested at 12K+ rows, 23 columns — completes in ~2 seconds

### Generic XML (everything else)

Recursive tree parser flattens XML into a flat table. Each direct child of the root element becomes one XLSX row.

- **Elements** → column paths (`parent.child`)
- **Attributes** → `@attr` columns
- **Repeated leaf values** → joined with `;`
- **Mixed content** → stored under `#text` key
- Column headers sorted alphabetically, auto-fitted width

## Install

```bash
go install github.com/davidyusaku-13/xml2xlsx@latest
```

Or build from source:

```bash
git clone https://github.com/davidyusaku-13/xml2xlsx.git
cd xml2xlsx
go build -o xml2xlsx .
```

## License

MIT
