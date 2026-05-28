package main

import (
	"encoding/xml"
	"strings"
	"testing"
)

func TestParseGenericXML_AndFlatten(t *testing.T) {
	inputXML := `<People>
  <Person id="1" active="true">
    <Name>Alice</Name>
    <Phone>111-222</Phone>
    <Phone>333-444</Phone>
    <Address>
      <City>New York</City>
      <Zip>10001</Zip>
    </Address>
  </Person>
  <Person id="2">
    <Name>Bob</Name>
    <Address>
      <City>Los Angeles</City>
    </Address>
  </Person>
</People>`

	dec := xml.NewDecoder(strings.NewReader(inputXML))
	rootStart, err := readFirstStartElement(dec)
	if err != nil {
		t.Fatalf("failed to read first start element: %v", err)
	}

	if isExcelXML(rootStart) {
		t.Fatalf("expected isExcelXML to be false")
	}

	root, err := parseGenericXML(dec, rootStart)
	if err != nil {
		t.Fatalf("failed to parse generic XML: %v", err)
	}

	if root.Name != "People" {
		t.Errorf("expected root name to be 'People', got %q", root.Name)
	}

	records := flattenRecords(root)
	if len(records) != 2 {
		t.Fatalf("expected 2 records, got %d", len(records))
	}

	// Verify Record 1 (Alice)
	r1 := records[0]
	expectedR1 := map[string]string{
		"@id":          "1",
		"@active":      "true",
		"Name":         "Alice",
		"Phone":        "111-222; 333-444",
		"Address.City": "New York",
		"Address.Zip":  "10001",
	}

	for k, expectedVal := range expectedR1 {
		gotVal, ok := r1[k]
		if !ok {
			t.Errorf("r1 missing expected key %q", k)
			continue
		}
		if gotVal != expectedVal {
			t.Errorf("r1[%q] = %q; want %q", k, gotVal, expectedVal)
		}
	}

	// Verify Record 2 (Bob)
	r2 := records[2-1]
	expectedR2 := map[string]string{
		"@id":          "2",
		"Name":         "Bob",
		"Address.City": "Los Angeles",
	}

	for k, expectedVal := range expectedR2 {
		gotVal, ok := r2[k]
		if !ok {
			t.Errorf("r2 missing expected key %q", k)
			continue
		}
		if gotVal != expectedVal {
			t.Errorf("r2[%q] = %q; want %q", k, gotVal, expectedVal)
		}
	}

	// Address.Zip should not exist in Bob's record
	if _, ok := r2["Address.Zip"]; ok {
		t.Errorf("r2 should not have key 'Address.Zip'")
	}

	// Test writeXLSX works
	tmpPath := t.TempDir() + "/generic.xlsx"
	err = writeXLSX(tmpPath, records)
	if err != nil {
		t.Fatalf("failed to write generic XLSX: %v", err)
	}
}
