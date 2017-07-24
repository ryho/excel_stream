package excel_stream

import (
	"bytes"
	"fmt"
	"io"
	"reflect"
	"testing"

	"github.com/tealeg/xlsx"
)

func TestXlsxStreamWrite(t *testing.T) {
	// When shouldMakeRealFiles is set to true this test will make actual XLSX files in the file system.
	// This is useful to ensure files open in Excel, Numbers, Google Docs, etc.
	// In case of issues you can use "Open XML SDK 2.5" to diagnose issues in generated XLSX files:
	// https://www.microsoft.com/en-us/download/details.aspx?id=30425
	shouldMakeRealFiles := false
	testCases := []struct {
		testName      string
		sheetNames    []string
		workbookData  [][][]string
		expectedError error
	}{
		{
			testName: "One Sheet",
			sheetNames: []string{
				"Sheet1",
			},
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300", "0000000123"},
				},
			},
		},
		{
			testName: "Several Sheets, with different numbers of columns and rows",
			sheetNames: []string{
				"Sheet 1", "Sheet 2", "Sheet3",
			},
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300", "0000000123"},
				},
				{
					{"Token", "Name", "Price", "SKU", "Stock"},
					{"456", "Salsa", "200", "0346", "1"},
					{"789", "Burritos", "400", "754", "3"},
				},
				{
					{"Token", "Name", "Price"},
					{"9853", "Guacamole", "500"},
					{"2357", "Margarita", "700"},
				},
			},
		},
		{
			testName: "One Sheet Registered, tries to write to two",
			sheetNames: []string{
				"Sheet 1",
			},
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300", "0000000123"},
				},
				{
					{"Token", "Name", "Price", "SKU"},
					{"456", "Salsa", "200", "0346"},
				},
			},
			expectedError: AlreadyOnLastSheetError,
		},
		{
			testName: "One Sheet, too many columns in row 1",
			sheetNames: []string{
				"Sheet 1",
			},
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300", "0000000123", "asdf"},
				},
			},
			expectedError: WrongNumberOfRowsError,
		},
		{
			testName: "One Sheet, too few columns in row 1",
			sheetNames: []string{
				"Sheet 1",
			},
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300"},
				},
			},
			expectedError: WrongNumberOfRowsError,
		},
		{
			testName: "Lots of Sheets, only writes rows to one, only writes headers to one, should not error and should still create a valid file",
			sheetNames: []string{
				"Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4", "Sheet 5", "Sheet 6",
			},
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300", "0000000123"},
				},
				{{}},
				{{"Id", "Unit Cost"}},
				{{}},
				{{}},
				{{}},
			},
		},
		{
			testName: "Two Sheets, only writes to one, should not error and should still create a valid file",
			sheetNames: []string{
				"Sheet 1", "Sheet 2",
			},
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300", "0000000123"},
				},
				{{}},
			},
		},
		{
			testName: "Larger Sheet",
			sheetNames: []string{
				"Sheet 1",
			},
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU", "Token", "Name", "Price", "SKU", "Token", "Name", "Price", "SKU", "Token", "Name", "Price", "SKU", "Token", "Name", "Price", "SKU", "Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123"},
					{"456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346"},
					{"789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754"},
					{"123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123"},
					{"456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346"},
					{"789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754"},
					{"123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123"},
					{"456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346"},
					{"789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754"},
					{"123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123"},
					{"456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346"},
					{"789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754"},
					{"123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123"},
					{"456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346"},
					{"789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754"},
					{"123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123"},
					{"456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346"},
					{"789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754"},
					{"123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123"},
					{"456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346"},
					{"789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754"},
					{"123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123"},
					{"456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346"},
					{"789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754"},
				},
			},
		},
		{
			testName: "UTF-8 Characters. This XLSX File loads correctly with Excel, Numbers, and Google Docs. It also passes Microsoft's Office File Format Validator.",
			sheetNames: []string{
				"Sheet1",
			},
			workbookData: [][][]string{
				{
					// String courtesy of https://github.com/minimaxir/big-list-of-naughty-strings/
					// Header row contains the tags that I am filtering on
					{"Token", endSheetDataTag, "Price", fmt.Sprintf(dimensionTag, "A1:D1")},
					// Japanese and emojis
					{"123", "パーティーへ行かないか", "300", "🍕🐵 🙈 🙉 🙊"},
					// XML encoder/parser test strings
					{"123", `<?xml version="1.0" encoding="ISO-8859-1"?>`, "300", `<?xml version="1.0" encoding="ISO-8859-1"?><!DOCTYPE foo [ <!ELEMENT foo ANY ><!ENTITY xxe SYSTEM "file:///etc/passwd" >]><foo>&xxe;</foo>`},
					// Upside down text and Right to Left Arabic text
					{"123", `˙ɐnbᴉlɐ ɐuƃɐɯ ǝɹolop ʇǝ ǝɹoqɐl ʇn ʇunpᴉpᴉɔuᴉ ɹodɯǝʇ poɯsnᴉǝ op pǝs 'ʇᴉlǝ ƃuᴉɔsᴉdᴉpɐ ɹnʇǝʇɔǝsuoɔ 'ʇǝɯɐ ʇᴉs ɹolop ɯnsdᴉ ɯǝɹo˥
					00˙Ɩ$-`, "300", `ﷺ`},
					{"123", "Taco", "300", "0000000123"},
				},
			},
		},
	}
	for i, testCase := range testCases {
		t.Run(testCase.testName, func(t *testing.T) {
			var filePath string
			var buffer *bytes.Buffer
			if shouldMakeRealFiles {
				filePath = fmt.Sprintf("Workbook%d.xlsx", i)
			} else {
				buffer = bytes.NewBuffer(nil)
			}
			err := writeStreamFile(filePath, buffer, testCase.sheetNames, testCase.workbookData, shouldMakeRealFiles)
			if err != testCase.expectedError {
				t.Fatalf("Error differs from expected error. Error: %v, Expected Error: %v ", err, testCase.expectedError)
			}
			if testCase.expectedError != nil {
				return
			}
			// read the file back with a 3rd party library
			var bufReader *bytes.Reader
			var size int64
			if !shouldMakeRealFiles {
				bufReader = bytes.NewReader(buffer.Bytes())
				size = bufReader.Size()
			}
			actualSheetNames, actualWorkbookData := readXLSXFile(t, filePath, bufReader, size, shouldMakeRealFiles)
			// check if data was able to be read correctly
			if !reflect.DeepEqual(actualSheetNames, testCase.sheetNames) {
				t.Fatal("Expected sheet names to be equal")
			}
			if !reflect.DeepEqual(actualWorkbookData, testCase.workbookData) {
				t.Fatal("Expected workbook data to be equal")
			}
		})
	}
}

// writeStreamFile will write the file using the Excel Stream library
func writeStreamFile(filePath string, fileBuffer io.Writer, sheetNames []string, workbookData [][][]string, shouldMakeRealFiles bool) error {
	var file *StreamFileBuilder
	var err error
	if shouldMakeRealFiles {
		file, err = NewStreamFileBuilderForPath(filePath)
		if err != nil {
			return err
		}
	} else {
		file = NewStreamFileBuilder(fileBuffer)
	}
	for i, sheetName := range sheetNames {
		header := workbookData[i][0]
		err := file.AddSheet(sheetName, header)
		if err != nil {
			return err
		}
	}
	excelStream, err := file.Build()
	if err != nil {
		return err
	}
	for i, sheetData := range workbookData {
		if i != 0 {
			err = excelStream.NextSheet()
			if err != nil {
				return err
			}
		}
		for i, row := range sheetData {
			if i == 0 {
				continue
			}
			err = excelStream.WriteRow(row)
			if err != nil {
				return err
			}
		}
	}
	err = excelStream.Close()
	if err != nil {
		return err
	}
	return nil
}

// readXLSXFile will read the file using the imported library github.com/tealeg/xlsx so that the written file can be checked.
func readXLSXFile(t *testing.T, filePath string, fileBuffer io.ReaderAt, size int64, shouldMakeRealFiles bool) ([]string, [][][]string) {
	var readFile *xlsx.File
	var err error
	if shouldMakeRealFiles {
		readFile, err = xlsx.OpenFile(filePath)
		if err != nil {
			t.Fatal(err)
		}
	} else {
		readFile, err = xlsx.OpenReaderAt(fileBuffer, size)
		if err != nil {
			t.Fatal(err)
		}
	}
	var actualWorkbookData [][][]string
	var sheetNames []string
	for _, sheet := range readFile.Sheets {
		sheetData := [][]string{}
		for _, row := range sheet.Rows {
			data := []string{}
			for _, cell := range row.Cells {
				str, err := cell.FormattedValue()
				if err != nil {
					t.Fatal(err)
				}
				data = append(data, str)
			}
			sheetData = append(sheetData, data)
		}
		sheetNames = append(sheetNames, sheet.Name)
		actualWorkbookData = append(actualWorkbookData, sheetData)
	}
	return sheetNames, actualWorkbookData
}
