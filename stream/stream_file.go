package excel_stream

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"io"
	"strconv"

	"github.com/tealeg/xlsx"
)

type StreamFile struct {
	xlsxFile       *xlsx.File
	sheetXmlPrefix []string
	sheetXmlSuffix []string
	zipWriter      *zip.Writer
	currentSheet   *streamSheet
}

type streamSheet struct {
	// sheetIndex is the Excel sheet index, which starts at 1
	index int
	// The number of rows that have been written to the sheet so far
	rowCount int
	// The number of columns in the sheet
	columnCount int
	// The writer to write to this sheet's file in the XLSX Zip file
	writer io.Writer
}

var (
	NoCurrentSheetError     = errors.New("No Current Sheet")
	WrongNumberOfRowsError  = errors.New("Invalid number of cells passed to WriteRow. All calls to WriteRow on the same sheet must have the same number of cells.")
	AlreadyOnLastSheetError = errors.New("NextSheet() called, but already on last sheet.")
	UnsupportedCellType     = errors.New("Unsupported cell type")
	UnknownCellType         = errors.New("Unknown cell type")
)

// WriteRow will write a row of cells to the current sheet. Every call to WriteRow on the same sheet must contain the
// same number of cells as the header provided when the sheet was created or an error will be returned. This function
// will always trigger a flush on success. Currently the only supported data type is string data.
func (sf *StreamFile) WriteRow(cells []string) error {
	if sf.currentSheet == nil {
		return NoCurrentSheetError
	}
	if len(cells) != sf.currentSheet.columnCount {
		return WrongNumberOfRowsError
	}
	sf.currentSheet.rowCount++
	if err := sf.currentSheet.write(`<row r="` + strconv.Itoa(sf.currentSheet.rowCount) + `">`); err != nil {
		return err
	}
	for colIndex, cellData := range cells {
		cellCoordinate := xlsx.GetCellIDStringFromCoords(colIndex, sf.currentSheet.rowCount-1)
		cellType, err := cellTypeString(xlsx.CellTypeInline)
		if err != nil {
			return err
		}

		cellOpen := `<c r="` + cellCoordinate + `" t="` + cellType + `"><is><t>`
		cellClose := `</t></is></c>`

		if err := sf.currentSheet.write(cellOpen); err != nil {
			return err
		}
		if err := xml.EscapeText(sf.currentSheet.writer, []byte(cellData)); err != nil {
			return err
		}
		if err := sf.currentSheet.write(cellClose); err != nil {
			return err
		}
	}
	if err := sf.currentSheet.write(`</row>`); err != nil {
		return err
	}
	return sf.zipWriter.Flush()
}

// NextSheet will switch to the next sheet. Sheets are selected in the same order they were added.
// Once you leave a sheet, you cannot return to it.
func (sf *StreamFile) NextSheet() error {
	var sheetIndex int
	if sf.currentSheet != nil {
		if sf.currentSheet.index >= len(sf.xlsxFile.Sheets) {
			return AlreadyOnLastSheetError
		}
		if err := sf.writeSheetEnd(); err != nil {
			sf.currentSheet = nil
			return err
		}
		sheetIndex = sf.currentSheet.index
	}
	sheetIndex++
	sf.currentSheet = &streamSheet{
		index:       sheetIndex,
		columnCount: len(sf.xlsxFile.Sheets[sheetIndex-1].Cols),
		rowCount:    1,
	}
	sheetPath := sheetFilePathPrefix + strconv.Itoa(sf.currentSheet.index) + sheetFilePathSuffix
	// There are two compression methods that the Golang zip.Writer supports, Store and Deflate, and we must use
	// Store here.
	// Deflate is one of the compression algorithms that .zip supports. Golang's implementation of Deflate will keep
	// everything passed to Write() and will only pass it down when Close() is called. Using this would prevent this
	// library from streaming with in an Excel sheet.
	// Store uses no compression and is just a no-op wrapper. Using this will allow data passed to WriteRow to get written
	// and then immediately flushed out to the network.
	fileWriter, err := sf.zipWriter.CreateHeader(&zip.FileHeader{Name: sheetPath, Method: zip.Store})
	if err != nil {
		return err
	}
	sf.currentSheet.writer = fileWriter

	if err := sf.writeSheetStart(); err != nil {
		return err
	}
	return nil
}

// Close closes the Stream File.
// Any sheets that have not yet been written to will have an empty sheet created for them.
func (sf *StreamFile) Close() error {
	// If there are sheets that have not been written yet, call NextSheet() which will add files to the zip for them.
	// XLSX readers may error if the sheets registered in the metadata are not present in the file.
	if sf.currentSheet != nil {
		for sf.currentSheet.index < len(sf.xlsxFile.Sheets) {
			if err := sf.NextSheet(); err != nil {
				return err
			}
		}
		// Write the end of the last sheet.
		if err := sf.writeSheetEnd(); err != nil {
			return err
		}
	}
	return sf.zipWriter.Close()
}

// cellTypeString returns the string value that should be used for the cell type.
// Unsupported or unknown cell types will return an error
// documentation for the c.t (cell.Type) attribute:
// b (Boolean): Cell containing a boolean.
// d (Date): Cell contains a date in the ISO 8601 format.
// e (Error): Cell containing an error.
// inlineStr (Inline String): Cell containing an (inline) rich string, i.e., one not in the shared string table.
// If this cell type is used, then the cell value is in the is element rather than the v element in the cell (c element).
// n (Number): Cell containing a number.
// s (Shared String): Cell containing a shared string.
// str (String): Cell containing a formula string.
func cellTypeString(enum xlsx.CellType) (string, error) {
	var cellTypeString string
	switch enum {
	case xlsx.CellTypeInline:
		cellTypeString = "inlineStr"
	case xlsx.CellTypeString:
		fallthrough
	case xlsx.CellTypeFormula:
		fallthrough
	case xlsx.CellTypeNumeric:
		fallthrough
	case xlsx.CellTypeBool:
		fallthrough
	case xlsx.CellTypeError:
		fallthrough
	case xlsx.CellTypeDate:
		fallthrough
	case xlsx.CellTypeGeneral:
		return "", UnsupportedCellType
	default:
		return "", UnknownCellType
	}
	return cellTypeString, nil
}

// writeSheetStart will write the start of the Sheet's XML as returned from the XMSX library.
func (sf *StreamFile) writeSheetStart() error {
	if sf.currentSheet == nil {
		return NoCurrentSheetError
	}
	return sf.currentSheet.write(sf.sheetXmlPrefix[sf.currentSheet.index-1])
}

// writeSheetEnd will write the end of the Sheet's XML as returned from the XMSX library.
func (sf *StreamFile) writeSheetEnd() error {
	if sf.currentSheet == nil {
		return NoCurrentSheetError
	}
	if err := sf.currentSheet.write(endSheetDataTag); err != nil {
		return err
	}
	return sf.currentSheet.write(sf.sheetXmlSuffix[sf.currentSheet.index-1])
}

func (ss *streamSheet) write(data string) error {
	_, err := ss.writer.Write([]byte(data))
	return err
}
