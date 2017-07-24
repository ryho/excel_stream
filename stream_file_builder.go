// Authors: Ryan Hollis (ryanh@)

// The purpose of the StreamFile library is to allow streamed writing of XLSX files.
// It relies heavily on the XLSX library (github.com/tealeg/xlsx).
// Directions:
// 1. Create a StreamFileBuilder with NewStreamFileBuilder() or NewStreamFileBuilderForPath().
// 2. Add the sheets and their first row of data by calling AddSheet().
// 3. Call Build() to get a StreamFile. Once built, all functions on the builder will return an error.
// 4. Write to the StreamFile with WriteRow(). Writes begin on the first sheet. New rows are always written and flushed
// to the io. All rows written to the same sheet must have the same number of cells as the header provided when the sheet
// was created or an error will be returned.
// 5. Call NextSheet() to proceed to the next sheet. Once NextSheet() is called, the previous sheet can not be edited.
// 6. Call Close() to finish.

// Future work suggestions:
// Currently the only supported cell type is string, since the main reason this library was written was to prevent
// strings from being interpreted as numbers. It would be nice to have support for numbers and money so that the exported
// files could better take advantage of Excel's features.
// All text is written with the same text style. Support for additional text styles could be added to highlight certain
// data in the file.
// The current default style uses fonts that are not on Macs by default so opening the XLSX files in Numbers causes a
// pop up that says there are missing fonts. The font could be changed to something that is usually found on Mac and PC.

package excel_stream

import (
	"archive/zip"
	"errors"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx"
)

type StreamFileBuilder struct {
	built     bool
	xlsxFile  *xlsx.File
	zipWriter *zip.Writer
}

const (
	sheetFilePathPrefix = "xl/worksheets/sheet"
	sheetFilePathSuffix = ".xml"
	endSheetDataTag     = "</sheetData>"
	dimensionTag        = `<dimension ref="%s"></dimension>`
)

var BuiltExcelStreamBuilderError = errors.New("StreamFileBuilder has already been built, functions may no longer be used")

// NewExcelBuilder creates an StreamFileBuilder that will write to the the provided io.writer
func NewStreamFileBuilder(writer io.Writer) *StreamFileBuilder {
	return &StreamFileBuilder{
		zipWriter: zip.NewWriter(writer),
		xlsxFile:  xlsx.NewFile(),
	}
}

// NewExcelBuilderForFile takes the name of an XLSX file and returns a builder for it.
// The file will be created if it does not exist, or truncated if it does.
func NewStreamFileBuilderForPath(path string) (*StreamFileBuilder, error) {
	file, err := os.Create(path)
	if err != nil {
		return nil, err
	}
	return NewStreamFileBuilder(file), nil
}

// AddSheet will add sheets with the given name with the provided headers. The headers cannot be edited later, and all
// rows written to the sheet must contain the same number of cells as the header. Sheet names must be unique, or an
// error will be thrown.
func (sb *StreamFileBuilder) AddSheet(name string, headers []string) error {
	if sb.built {
		return BuiltExcelStreamBuilderError
	}
	sheet, err := sb.xlsxFile.AddSheet(name)
	if err != nil {
		// Set built on error so that all subsequent calls to the builder will also fail.
		sb.built = true
		return err
	}
	row := sheet.AddRow()
	if count := row.WriteSlice(&headers, -1); count != len(headers) {
		// Set built on error so that all subsequent calls to the builder will also fail.
		sb.built = true
		return errors.New("Failed to write headers")
	}
	return nil
}

// Build begins streaming the XLSX file to the io, by writing all the Excel metadata. It creates a StreamFile struct
// that can be used to write the rows to the sheets.
func (sb *StreamFileBuilder) Build() (*StreamFile, error) {
	if sb.built {
		return nil, BuiltExcelStreamBuilderError
	}
	sb.built = true
	parts, err := sb.xlsxFile.MarshallParts()
	if err != nil {
		return nil, err
	}
	es := &StreamFile{
		zipWriter:      sb.zipWriter,
		xlsxFile:       sb.xlsxFile,
		sheetXmlPrefix: make([]string, len(sb.xlsxFile.Sheets)),
		sheetXmlSuffix: make([]string, len(sb.xlsxFile.Sheets)),
	}
	for path, data := range parts {
		// If the part is a sheet, don't write it yet. We only want to write the Excel metadata files, since at this
		// point the sheets are still empty. The sheet files will be written later as their rows come in.
		if strings.HasPrefix(path, sheetFilePathPrefix) {
			if err := sb.processEmptySheetXML(es, path, data); err != nil {
				return nil, err
			}
			continue
		}
		metadataFile, err := sb.zipWriter.Create(path)
		if err != nil {
			return nil, err
		}
		_, err = metadataFile.Write([]byte(data))
		if err != nil {
			return nil, err
		}
	}

	if err := es.NextSheet(); err != nil {
		return nil, err
	}
	return es, nil
}

// processEmptySheetXML will take in the path and XML data of an empty sheet, and will save the beginning and end of the
// XML file so that these can be written at the right time.
func (sb *StreamFileBuilder) processEmptySheetXML(sf *StreamFile, path, data string) error {
	// Get the sheet index from the path
	sheetIndex, err := getSheetIndex(sf, path)
	if err != nil {
		return err
	}

	// Remove the Dimension tag. Since more rows are going to be written to the sheet, it will be wrong.
	// It is valid to for a sheet to be missing a Dimension tag, but it is not valid for it to be wrong.
	data, err = removeDimensionTag(data, sf.xlsxFile.Sheets[sheetIndex])
	if err != nil {
		return err
	}

	// Split the sheet at the end of its SheetData tag so that more rows can be added inside.
	prefix, suffix, err := splitSheetIntoPrefixAndSuffix(data)
	if err != nil {
		return err
	}
	sf.sheetXmlPrefix[sheetIndex] = prefix
	sf.sheetXmlSuffix[sheetIndex] = suffix
	return nil
}

// getSheetIndex parses the path to the Excel sheet data and returns the index
// The files that store the data for each sheet must have the format:
// xl/worksheets/sheet123.xml
// where 123 is the index of the sheet. This file path format is part of the XLSX file standard.
func getSheetIndex(sf *StreamFile, path string) (int, error) {
	indexString := path[len(sheetFilePathPrefix) : len(path)-len(sheetFilePathSuffix)]
	sheetExcelIndex, err := strconv.Atoi(indexString)
	if err != nil {
		return -1, errors.New("Unexpected sheet file name from XLSX library")
	}
	if sheetExcelIndex < 1 || len(sf.sheetXmlPrefix) < sheetExcelIndex ||
		len(sf.sheetXmlSuffix) < sheetExcelIndex || len(sf.xlsxFile.Sheets) < sheetExcelIndex {
		return -1, errors.New("Unexpected sheet index")
	}
	sheetArrayIndex := sheetExcelIndex - 1
	return sheetArrayIndex, nil
}

// removeDimensionTag will return the passed in Excel Spreadsheet XML with the dimension tag removed.
// data is the XML data for the sheet
// sheet is the xlsx.Sheet struct that the XML was created from.
// Can return an error if the XML's dimension tag does not match was is expected based on the provided Sheet
func removeDimensionTag(data string, sheet *xlsx.Sheet) (string, error) {
	x := len(sheet.Cols) - 1
	y := len(sheet.Rows) - 1
	var dimensionRef string
	if x < 0 || y < 0 {
		dimensionRef = "A1"
	} else {
		endCoordinate := xlsx.GetCellIDStringFromCoords(x, y)
		dimensionRef = "A1:" + endCoordinate
	}
	dataParts := strings.Split(data, fmt.Sprintf(dimensionTag, dimensionRef))
	if len(dataParts) != 2 {
		return "", errors.New("Unexpected Sheet XML from XLSX library. Dimension tag not found.")
	}
	return dataParts[0] + dataParts[1], nil
}

// splitSheetIntoPrefixAndSuffix will split the provided XML sheet into a prefix and a suffix so that
// more Excel rows can be inserted in between.
func splitSheetIntoPrefixAndSuffix(data string) (string, string, error) {
	// Split the sheet at the end of its SheetData tag so that more rows can be added inside.
	sheetParts := strings.Split(data, endSheetDataTag)
	if len(sheetParts) != 2 {
		return "", "", errors.New("Unexpected Sheet XML from XLSX library. SheetData close tag not found.")
	}
	return sheetParts[0], sheetParts[1], nil
}
