package xlmapper

import (
	"errors"
	"github.com/dizk/xlsx"
	"reflect"
	"strconv"
)

const defaultStructTag = "xlmapper"

type XlsxDecoder struct {
	headers   []string
	data      [][]string
	structTag string
}

// NewXlsxDecoder creates a NewXlsxDecoder
func NewXlsxDecoder(filePath string, headerRow int, sheet int) (*XlsxDecoder, error) {
	if headerRow < 0 {
		return nil, errors.New("HeaderRow must be 0 or greater")
	}

	if sheet < 0 {
		return nil, errors.New("Sheet must be 0 or greater")
	}

	xs, err := xlsx.FileToSlice(filePath)
	if err != nil {
		return nil, err
	}

	if sheet >= len(xs) {
		return nil, errors.New("Invalid sheet index")
	}

	if headerRow+1 >= len(xs[sheet]) {
		return nil, errors.New("Not enough rows, must be at least header + 1")
	}

	xlsxDecoder := XlsxDecoder{
		structTag: defaultStructTag,
		headers:   xs[sheet][headerRow],
		data:      xs[sheet][headerRow+1:],
	}

	return &xlsxDecoder, err
}

func (x *XlsxDecoder) HasNextRow() bool {
	return len(x.data) > 0
}

func (x *XlsxDecoder) NextRow() (map[string]string, error) {
	if !x.HasNextRow() {
		return nil, errors.New("No more rows to get.")
	}

	var row []string
	row, x.data = x.data[0], x.data[1:]

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
	return x.unmarshallRow(row, v)
}

func (x *XlsxDecoder) unmarshallRow(row map[string]string, v interface{}) error {

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
		xlheader := fieldType.Tag.Get(x.structTag)
		if xlheader == "" {
			continue
		}

		rowValue := row[xlheader]
		if rowValue == "" {
			continue
		}

		// Getting the field in v
		field := vs.Field(i)
		err := setField(field, rowValue)
		if err != nil {
			return err
		}
	}

	return nil
}

func setField(field reflect.Value, value string) error {
	if !field.CanSet() {
		return errors.New("Unable to set value " + value + ".")
	}

	switch field.Kind() {
	case reflect.Int:
		i, err := strconv.Atoi(value)
		if err != nil {
			return err
		}
		field.SetInt(int64(i))
		return nil
	case reflect.String:
		field.SetString(value)
		return nil
	default:
		return errors.New("Invalid field type for " + value + ", must be int or string.")
	}
}
