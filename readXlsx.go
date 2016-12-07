package xlmapper

import (
	"errors"
	"github.com/tealeg/xlsx"
	"reflect"
	"strconv"
)

type XlsxDecoder struct {
	file       *xlsx.File
	sheet      int
	headerRow  int
	skipCells  int
	currentRow int
	headers    []string
	nextRow    []string
	finished   bool
}

// NewXlsxDecoder creates a NewXlsxDecoder
func NewXlsxDecoder(filePath string, headerRow int, sheet int) (*XlsxDecoder, error) {
	if headerRow < 0 {
		return nil, errors.New("HeaderRow must be 0 or greater")
	}

	if sheet < 0 {
		return nil, errors.New("Sheet must be 0 or greater")
	}

	xlsxDecoder := XlsxDecoder{
		sheet:      sheet,
		headerRow:  headerRow,
		skipCells:  0,
		currentRow: 0,
	}

	var err error

	xlsxDecoder.file, err = xlsx.OpenFile(filePath)
	if err != nil {
		return nil, err
	}

	if sheet >= len(xlsxDecoder.file.Sheets) {
		return nil, errors.New("Invalid sheet index")
	}

	if len(xlsxDecoder.file.Sheets[sheet].Rows) <= 1 {
		return nil, errors.New("Not enough rows, must be at least header + 1")
	}

	err = xlsxDecoder.readHeader()
	return &xlsxDecoder, err
}

func (x *XlsxDecoder) readHeader() error {
	hRow := x.file.Sheets[x.sheet].Rows[x.headerRow]
	doneSkipping := false

	for _, cell := range hRow.Cells {
		if cell.Value == "" && !doneSkipping {
			x.skipCells++
			continue
		}

		if cell.Value == "" && doneSkipping {
			break
		}

		doneSkipping = true

		x.headers = append(x.headers, cell.Value)
	}

	x.currentRow = x.headerRow + 1
	x.readNextRow()

	return nil
}

func (x *XlsxDecoder) HasNextRow() bool {
	return !x.finished
}

func (x *XlsxDecoder) NextRow() (map[string]string, error) {
	row := x.nextRow
	x.readNextRow()
	if x.nextRow == nil {
		x.finished = true
	}

	// Create a map with HEADER->VALUE
	headerToValue := make(map[string]string)
	for i := 0; i < len(x.headers); i++ {
		headerToValue[x.headers[i]] = row[i]
	}

	return headerToValue, nil
}

func (x *XlsxDecoder) UnmarshallNextRow(v interface{}) error {
	row, err := x.NextRow()
	if err != nil {
		return err
	}
	return x.UnmarshallRow(row, v)
}

func (x *XlsxDecoder) UnmarshallRow(row map[string]string, v interface{}) error {

	// Ensure v is a pointer
	vv := reflect.ValueOf(v)
	if vv.Kind() != reflect.Ptr {
		return errors.New("Invalid type, must be pointer to a struct.")
	}

	// Ensure v points to a struct
	vs := vv.Elem()
	if vs.Kind() != reflect.Struct {
		return errors.New("Invalid type, must be pointer to a struct.")
	}

	// Get the type of v
	vt := reflect.TypeOf(v).Elem()
	if vt == nil {
		return errors.New("Invalid type, must be pointer to a struct.")
	}

	// Iterate over the fields
	for i := 0; i < vt.NumField(); i++ {
		// Getting the xlheader field Tag
		fieldType := vt.Field(i)
		xlheader := fieldType.Tag.Get("xlheader")
		if xlheader == "" {
			continue
		}

		rv := row[xlheader]
		if rv == "" {
			continue
		}

		// Getting the field in v
		field := vs.Field(i)
		if field.CanSet() {

			switch field.Kind() {
			case reflect.Int:
				i, err := strconv.Atoi(rv)
				if err != nil {
					return err
				}
				field.SetInt(int64(i))
			case reflect.String:
				field.SetString(rv)
			default:
				return errors.New("Invalid field type for " + xlheader + ", must be int or string.")
			}
		} else {
			return errors.New("Unable to set value for " + xlheader + ".")
		}
	}

	return nil
}

func (x *XlsxDecoder) readNextRow() {
	if x.currentRow >= len(x.file.Sheets[x.sheet].Rows) {
		x.nextRow = nil
		return
	}

	cRow := x.file.Sheets[x.sheet].Rows[x.currentRow]
	x.currentRow++

	var res []string
	empty := true

	for i, cell := range cRow.Cells {
		if i < x.skipCells {
			continue
		}

		if i > len(x.headers) {
			break
		}

		if cell.Value != "" {
			empty = false
		}

		res = append(res, cell.Value)

	}

	if empty {
		res = nil
	}

	x.nextRow = res
}
