package main

import (
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// excelSheet represents one parsed worksheet from Excel XML Spreadsheet format.
type excelSheet struct {
	Name string
	Rows [][]interface{}
}

// isExcelXML checks if the root element is an Excel XML Spreadsheet workbook.
func isExcelXML(root xml.StartElement) bool {
	if root.Name.Local != "Workbook" {
		return false
	}
	for _, a := range root.Attr {
		if a.Name.Local == "xmlns" && a.Value == "urn:schemas-microsoft-com:office:spreadsheet" {
			return true
		}
	}
	return false
}

// readFirstStartElement reads tokens until the first StartElement.
func readFirstStartElement(dec *xml.Decoder) (xml.StartElement, error) {
	for {
		tok, err := dec.Token()
		if err != nil {
			return xml.StartElement{}, err
		}
		if start, ok := tok.(xml.StartElement); ok {
			return start, nil
		}
	}
}

// parseExcelXML parses a Microsoft Excel XML Spreadsheet workbook into sheets.
func parseExcelXML(dec *xml.Decoder, rootStart xml.StartElement) ([]excelSheet, error) {
	var sheets []excelSheet

	for {
		tok, err := dec.Token()
		if err != nil {
			if err == io.EOF {
				break
			}
			return nil, fmt.Errorf("reading Excel XML: %w", err)
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "Worksheet" {
				s, err := parseWorksheet(dec, t)
				if err != nil {
					return nil, err
				}
				sheets = append(sheets, s)
			} else {
				skipElement(dec)
			}
		case xml.EndElement:
			return sheets, nil
		}
	}

	return sheets, nil
}

// parseWorksheet reads a <Worksheet> element.
func parseWorksheet(dec *xml.Decoder, start xml.StartElement) (excelSheet, error) {
	sheet := excelSheet{
		Name: getAttr(start.Attr, "Name"),
	}
	if sheet.Name == "" {
		sheet.Name = "Sheet1"
	}

	for {
		tok, err := dec.Token()
		if err != nil {
			return sheet, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "Table" {
				rows, err := parseTable(dec)
				if err != nil {
					return sheet, err
				}
				sheet.Rows = rows
			} else {
				skipElement(dec)
			}
		case xml.EndElement:
			return sheet, nil
		}
	}
}

// parseTable reads a <Table> element and returns all rows.
func parseTable(dec *xml.Decoder) ([][]interface{}, error) {
	var rows [][]interface{}

	for {
		tok, err := dec.Token()
		if err != nil {
			return rows, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "Row" {
				row, err := parseRow(dec)
				if err != nil {
					return rows, err
				}
				rows = append(rows, row)
			} else {
				skipElement(dec)
			}
		case xml.EndElement:
			return rows, nil
		}
	}
}

// parseRow reads a <Row> element and returns its cell values.
func parseRow(dec *xml.Decoder) ([]interface{}, error) {
	var cells []interface{}
	for {
		tok, err := dec.Token()
		if err != nil {
			return cells, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "Cell" {
				val, err := parseCell(dec)
				if err != nil {
					return cells, err
				}
				cells = append(cells, val)
			} else {
				skipElement(dec)
			}
		case xml.EndElement:
			return cells, nil
		}
	}
}

// parseCell reads a <Cell> element and returns its Data value (or empty string).
// Numeric values (ss:Type="Number") are returned as float64; everything else as string.
func parseCell(dec *xml.Decoder) (interface{}, error) {
	var value interface{} = ""
	for {
		tok, err := dec.Token()
		if err != nil {
			return "", err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "Data" {
				text, err := parseData(dec)
				if err != nil {
					return "", err
				}
				dataType := getAttr(t.Attr, "Type")
				if dataType == "Number" {
					if f, err := strconv.ParseFloat(text, 64); err == nil {
						value = f
					} else {
						value = text
					}
				} else {
					value = text
				}
			} else {
				skipElement(dec)
			}
		case xml.EndElement:
			return value, nil
		}
	}
}

// parseData reads a <Data> element and returns its text content.
func parseData(dec *xml.Decoder) (string, error) {
	var buf strings.Builder
	for {
		tok, err := dec.Token()
		if err != nil {
			return buf.String(), err
		}

		switch t := tok.(type) {
		case xml.CharData:
			buf.Write(t)
		case xml.EndElement:
			return strings.TrimSpace(buf.String()), nil
		}
	}
}

// skipElement skips over one XML element (including all its children).
func skipElement(dec *xml.Decoder) error {
	depth := 1
	for {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
			if depth == 0 {
				return nil
			}
		}
	}
}

// getAttr returns an attribute value by its local name.
func getAttr(attrs []xml.Attr, localName string) string {
	for _, a := range attrs {
		if a.Name.Local == localName {
			return a.Value
		}
	}
	return ""
}

// writeExcelXLSX writes parsed Excel XML sheets to an XLSX file using stream writer.
func writeExcelXLSX(path string, sheets []excelSheet) error {
	f := excelize.NewFile()
	defer f.Close()

	for i, sheet := range sheets {
		var sheetName string
		if i == 0 {
			sheetName = "Sheet1"
			f.SetSheetName(sheetName, sheet.Name)
			sheetName = sheet.Name
		} else {
			idx, err := f.NewSheet(sheet.Name)
			if err != nil {
				return fmt.Errorf("new sheet: %w", err)
			}
			sheetName = f.GetSheetName(idx)
		}
		if len(sheet.Rows) == 0 {
			continue
		}

		sw, err := f.NewStreamWriter(sheetName)
		if err != nil {
			return fmt.Errorf("stream writer for %q: %w", sheetName, err)
		}

		for rowIdx, row := range sheet.Rows {
			cells := make([]interface{}, len(row))
			for j, val := range row {
				cells[j] = val
			}
			axis, err := excelize.CoordinatesToCellName(1, rowIdx+1)
			if err != nil {
				return fmt.Errorf("axis: %w", err)
			}
			if err := sw.SetRow(axis, cells); err != nil {
				return fmt.Errorf("set row %d: %w", rowIdx+1, err)
			}
		}

		if err := sw.Flush(); err != nil {
			return fmt.Errorf("flush: %w", err)
		}
	}

	return f.SaveAs(path)
}
