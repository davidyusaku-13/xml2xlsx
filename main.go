package main

import (
	"encoding/xml"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

func main() {
	var input, output string

	switch len(os.Args) {
	case 2:
		input = os.Args[1]
		ext := filepath.Ext(input)
		if strings.ToLower(ext) == ".xml" {
			output = strings.TrimSuffix(input, ext) + ".xlsx"
		} else {
			output = input + ".xlsx"
		}
	case 3:
		input = os.Args[1]
		output = os.Args[2]
	default:
		fmt.Fprintf(os.Stderr, "Usage: xml2xlsx <input.xml> [output.xlsx]\n")
		os.Exit(1)
	}

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

	var errConvert error
	if isExcelXML(rootStart) {
		errConvert = convertExcelXML(dec, rootStart, input, output)
	} else {
		errConvert = convertGeneric(dec, rootStart, input, output)
	}

	if errConvert != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", errConvert)
		os.Exit(1)
	}
}

func convertExcelXML(dec *xml.Decoder, rootStart xml.StartElement, input, output string) error {
	sheets, err := parseExcelXML(dec, rootStart)
	if err != nil {
		return fmt.Errorf("parsing Excel XML: %w", err)
	}
	if len(sheets) == 0 {
		return fmt.Errorf("no worksheets found")
	}
	if err := writeExcelXLSX(output, sheets); err != nil {
		return fmt.Errorf("writing XLSX: %w", err)
	}
	totalRows := 0
	for _, s := range sheets {
		totalRows += len(s.Rows)
	}
	fmt.Printf("Converted %s → %s (%d sheets, %d rows)\n", input, output, len(sheets), totalRows)
	return nil
}

func convertGeneric(dec *xml.Decoder, rootStart xml.StartElement, input, output string) error {
	root, err := parseGenericXML(dec, rootStart)
	if err != nil {
		return fmt.Errorf("parsing XML: %w", err)
	}
	records := flattenRecords(root)
	if len(records) == 0 {
		return fmt.Errorf("no records found (root element has no child elements)")
	}
	if err := writeXLSX(output, records); err != nil {
		return fmt.Errorf("writing XLSX: %w", err)
	}
	keySet := make(map[string]struct{})
	for _, rec := range records {
		for k := range rec {
			keySet[k] = struct{}{}
		}
	}
	fmt.Printf("Converted %s → %s (%d rows, %d columns)\n", input, output, len(records), len(keySet))
	return nil
}
