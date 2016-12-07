package xlmapper

import "testing"

const excelFileName = "testdata/simple_sheet.xlsx"
const excelFileName2 = "testdata/testFileToSlice.xlsx"

type SimpleSheetLine struct {
	Number int    `xlmapper:"Number"`
	Name   string `xlmapper:"Name"`
	Field  string `xlmapper:"Field"`
}

func TestNewDecoder(t *testing.T) {
	_, err := NewXlsxDecoder(excelFileName, 0, 0)

	if err != nil {
		t.Error("Expected err to be nil, was ", err)
	}

	_, err = NewXlsxDecoder(excelFileName2, 0, 0)

	if err != nil {
		t.Error("Expected err to be nil, was ", err)
	}
}

func TestHasNextRow(t *testing.T) {
	xd, _ := NewXlsxDecoder(excelFileName, 0, 0)

	v := xd.HasNextRow()

	if v != true {
		t.Error("Expected true, was ", v)
	}
}

func TestUnMarshallRow(t *testing.T) {
	xd, _ := NewXlsxDecoder(excelFileName, 0, 0)

	var v SimpleSheetLine

	err := xd.UnmarshallNextRow(&v)

	if err != nil {
		t.Error("Expected err to be nil, was ", err)
	}

	if v.Name != "Ole" {
		t.Error("Expected Ole, was ", v.Name)
	}

	if v.Number != 1 {
		t.Error("Expected 1, was ", v.Number)
	}

	if v.Field != "One" {
		t.Error("Expected One, was ", v.Field, " header was ", xd.headers)

	}

}
