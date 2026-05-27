package main

import (
	"encoding/xml"
	"fmt"
	"sort"
	"strings"

	"github.com/xuri/excelize/v2"
)

// xmlNode is a generic XML element tree node.
type xmlNode struct {
	Name     string
	Attrs    map[string]string
	Children []*xmlNode
	Text     string
}

// parseGenericXML reads an XML document from decoder, given the root element.
func parseGenericXML(dec *xml.Decoder, rootStart xml.StartElement) (*xmlNode, error) {
	root := &xmlNode{Name: rootStart.Name.Local}
	for _, a := range rootStart.Attr {
		if root.Attrs == nil {
			root.Attrs = make(map[string]string, len(rootStart.Attr))
		}
		root.Attrs[a.Name.Local] = a.Value
	}
	if err := decodeChildren(dec, root); err != nil {
		return nil, err
	}
	return root, nil
}

// decodeChildren reads child tokens until the matching EndElement,
// populating node.Children and node.Text.
func decodeChildren(dec *xml.Decoder, node *xmlNode) error {
	for {
		tok, err := dec.Token()
		if err != nil {
			return fmt.Errorf("reading XML: %w", err)
		}

		switch t := tok.(type) {
		case xml.StartElement:
			child := &xmlNode{Name: t.Name.Local}
			for _, a := range t.Attr {
				if child.Attrs == nil {
					child.Attrs = make(map[string]string, len(t.Attr))
				}
				child.Attrs[a.Name.Local] = a.Value
			}
			if err := decodeChildren(dec, child); err != nil {
				return err
			}
			node.Children = append(node.Children, child)

		case xml.CharData:
			text := strings.TrimSpace(string(t))
			if text != "" {
				node.Text += text
			}

		case xml.EndElement:
			return nil
		}
	}
}

// flattenRecords converts root's direct children into flat string-keyed records.
func flattenRecords(root *xmlNode) []map[string]string {
	if root == nil {
		return nil
	}
	records := make([]map[string]string, 0, len(root.Children))
	for _, child := range root.Children {
		rec := make(map[string]string)
		flattenNode(child, "", rec)
		records = append(records, rec)
	}
	return records
}

func flattenNode(node *xmlNode, prefix string, dest map[string]string) {
	// Attributes at the current level
	for k, v := range node.Attrs {
		key := prefix + "@" + k
		if prev, ok := dest[key]; ok {
			dest[key] = prev + "; " + v
		} else {
			dest[key] = v
		}
	}

	if len(node.Children) > 0 {
		if node.Text != "" {
			dest[prefix+"#text"] = node.Text
		}
		for _, child := range node.Children {
			flattenNode(child, prefix+child.Name+".", dest)
		}
	} else if node.Text != "" {
		col := prefix
		if col == "" {
			col = node.Name
		} else {
			col = strings.TrimSuffix(col, ".")
		}
		if prev, ok := dest[col]; ok {
			dest[col] = prev + "; " + node.Text
		} else {
			dest[col] = node.Text
		}
	}
}

// writeXLSX writes generic flattened records to a new XLSX file.
func writeXLSX(path string, records []map[string]string) error {
	f := excelize.NewFile()
	defer f.Close()

	keySet := make(map[string]struct{})
	for _, rec := range records {
		for k := range rec {
			keySet[k] = struct{}{}
		}
	}
	keys := make([]string, 0, len(keySet))
	for k := range keySet {
		keys = append(keys, k)
	}
	sort.Strings(keys)

	sheet := "Sheet1"
	const headerRow = 1

	for i, key := range keys {
		cell, err := excelize.CoordinatesToCellName(i+1, headerRow)
		if err != nil {
			return fmt.Errorf("cell coordinate: %w", err)
		}
		if err := f.SetCellStr(sheet, cell, key); err != nil {
			return fmt.Errorf("set cell: %w", err)
		}
	}

	for rowIdx, rec := range records {
		for colIdx, key := range keys {
			val, ok := rec[key]
			if !ok {
				continue
			}
			cell, err := excelize.CoordinatesToCellName(colIdx+1, rowIdx+2)
			if err != nil {
				return fmt.Errorf("cell coordinate: %w", err)
			}
			if err := f.SetCellStr(sheet, cell, val); err != nil {
				return fmt.Errorf("set cell: %w", err)
			}
		}
	}

	for i, key := range keys {
		colLetter, err := excelize.ColumnNumberToName(i + 1)
		if err != nil {
			continue
		}
		maxW := float64(len(key))
		for _, rec := range records {
			if val, ok := rec[key]; ok && float64(len(val)) > maxW {
				maxW = float64(len(val))
			}
		}
		width := maxW + 2
		if width > 50 {
			width = 50
		}
		_ = f.SetColWidth(sheet, colLetter, colLetter, width)
	}

	return f.SaveAs(path)
}
