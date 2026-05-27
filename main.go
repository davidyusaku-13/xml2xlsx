package main

import (
	"encoding/xml"
	"fmt"
	"os"
)

func main() {
	if len(os.Args) != 3 {
		fmt.Fprintf(os.Stderr, "Usage: xml2xlsx <input.xml> <output.xlsx>\n")
		os.Exit(1)
	}

	input := os.Args[1]
	output := os.Args[2]

	f, err := os.Open(input)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: cannot open %s: %v\n", input, err)
		os.Exit(1)
	}
	defer f.Close()

	dec := xml.NewDecoder(f)
	rootStart, err := readFirstStartElement(dec)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error reading XML: %v\n", err)
		os.Exit(1)
	}

	if isExcelXML(rootStart) {
		convertExcelXML(dec, rootStart, input, output)
	} else {
		convertGeneric(dec, rootStart, input, output)
	}
}

func convertExcelXML(dec *xml.Decoder, rootStart xml.StartElement, input, output string) {
	sheets, err := parseExcelXML(dec, rootStart)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error parsing Excel XML: %v\n", err)
		os.Exit(1)
	}
	if len(sheets) == 0 {
		fmt.Fprintf(os.Stderr, "Error: no worksheets found\n")
		os.Exit(1)
	}
	if err := writeExcelXLSX(output, sheets); err != nil {
		fmt.Fprintf(os.Stderr, "Error writing XLSX: %v\n", err)
		os.Exit(1)
	}
	totalRows := 0
	for _, s := range sheets {
		totalRows += len(s.Rows)
	}
	fmt.Printf("Converted %s → %s (%d sheets, %d rows)\n", input, output, len(sheets), totalRows)
}

func convertGeneric(dec *xml.Decoder, rootStart xml.StartElement, input, output string) {
	root, err := parseGenericXML(dec, rootStart)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error parsing XML: %v\n", err)
		os.Exit(1)
	}
	records := flattenRecords(root)
	if len(records) == 0 {
		fmt.Fprintf(os.Stderr, "Error: no records found (root element has no child elements)\n")
		os.Exit(1)
	}
	if err := writeXLSX(output, records); err != nil {
		fmt.Fprintf(os.Stderr, "Error writing XLSX: %v\n", err)
		os.Exit(1)
	}
	keySet := make(map[string]struct{})
	for _, rec := range records {
		for k := range rec {
			keySet[k] = struct{}{}
		}
	}
	fmt.Printf("Converted %s → %s (%d rows, %d columns)\n", input, output, len(records), len(keySet))
}
