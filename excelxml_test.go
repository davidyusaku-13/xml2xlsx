package main

import (
	"encoding/xml"
	"strings"
	"testing"
)

func TestSanitizeSheetName(t *testing.T) {
	tests := []struct {
		input    string
		expected string
	}{
		{"NormalName", "NormalName"},
		{"Sheet [Name] : Invalid * ?", "Sheet Name  Invalid  "},
		{"Very Long Sheet Name That Exceeds Thirty One Characters Limit", "Very Long Sheet Name That Excee"},
		{"", "Sheet"},
		{"\\\\/?*:[]", "Sheet"}, // all invalid characters -> empty -> "Sheet"
	}

	for _, tc := range tests {
		got := sanitizeSheetName(tc.input)
		if got != tc.expected {
			t.Errorf("sanitizeSheetName(%q) = %q; want %q", tc.input, got, tc.expected)
		}
	}
}

func TestParseExcelXML_SparseAndSanitization(t *testing.T) {
	inputXML := `<?xml version="1.0"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
          xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
  <Worksheet ss:Name="Test / Sheet * Name">
    <Table>
      <Row ss:Index="2">
        <Cell ss:Index="2"><Data ss:Type="String">B2 Value</Data></Cell>
        <Cell ss:Index="4"><Data ss:Type="Number">123.45</Data></Cell>
      </Row>
      <Row>
        <Cell><Data ss:Type="String">A3 Value</Data></Cell>
      </Row>
    </Table>
  </Worksheet>
  <Worksheet ss:Name="Test / Sheet * Name">
    <Table>
      <Row>
        <Cell><Data ss:Type="String">Sheet 2 Row 1</Data></Cell>
      </Row>
    </Table>
  </Worksheet>
</Workbook>`

	dec := xml.NewDecoder(strings.NewReader(inputXML))
	rootStart, err := readFirstStartElement(dec)
	if err != nil {
		t.Fatalf("failed to read first start element: %v", err)
	}

	if !isExcelXML(rootStart) {
		t.Fatalf("expected isExcelXML to be true")
	}

	sheets, err := parseExcelXML(dec, rootStart)
	if err != nil {
		t.Fatalf("failed to parse Excel XML: %v", err)
	}

	if len(sheets) != 2 {
		t.Fatalf("expected 2 sheets, got %d", len(sheets))
	}

	// Verify Sheet 1
	s1 := sheets[0]
	if s1.Name != "Test / Sheet * Name" {
		t.Errorf("expected original name 'Test / Sheet * Name', got %q", s1.Name)
	}

	// Row 0 should be nil (due to Row index=2)
	if len(s1.Rows) != 3 {
		t.Fatalf("expected 3 rows in sheet 1, got %d", len(s1.Rows))
	}
	if s1.Rows[0] != nil {
		t.Errorf("expected row 0 to be nil (empty/skipped), got %v", s1.Rows[0])
	}

	// Row 1 (which represents Row 2) should have size 4: ["", "B2 Value", "", 123.45]
	r1 := s1.Rows[1]
	if len(r1) != 4 {
		t.Fatalf("expected row 1 to have 4 cells, got %d: %v", len(r1), r1)
	}
	if r1[0] != "" || r1[1] != "B2 Value" || r1[2] != "" || r1[3] != 123.45 {
		t.Errorf("unexpected cells in row 1: %v", r1)
	}

	// Row 2 (which represents Row 3) should have size 1: ["A3 Value"]
	r2 := s1.Rows[2]
	if len(r2) != 1 {
		t.Fatalf("expected row 2 to have 1 cell, got %d: %v", len(r2), r2)
	}
	if r2[0] != "A3 Value" {
		t.Errorf("expected cell 0 in row 2 to be 'A3 Value', got %v", r2[0])
	}

	// Verify writeExcelXLSX behavior with unique sheet name generation
	// We can write to a temporary file
	tmpPath := t.TempDir() + "/output.xlsx"
	err = writeExcelXLSX(tmpPath, sheets)
	if err != nil {
		t.Fatalf("failed to write XLSX: %v", err)
	}
}
