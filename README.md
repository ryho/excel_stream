# excel_stream
The purpose of the StreamFile library is to allow streamed writing of XLSX files.
It relies heavily on the XLSX library (github.com/tealeg/xlsx).
Directions:
1. Create a StreamFileBuilder with NewStreamFileBuilder() or NewStreamFileBuilderForPath().
2. Add the sheets and their first row of data by calling AddSheet().
3. Call Build() to get a StreamFile. Once built, all functions on the builder will return an error.
4. Write to the StreamFile with WriteRow(). Writes begin on the first sheet. New rows are always written and flushed
to the io. All rows written to the same sheet must have the same number of cells as the header provided when the sheet
was created or an error will be returned.
5. Call NextSheet() to proceed to the next sheet. Once NextSheet() is called, the previous sheet can not be edited.
6. Call Close() to finish.

Future work suggestions:
Currently the only supported cell type is string, since the main reason this library was written was to prevent
strings from being interpreted as numbers. It would be nice to have support for numbers and money so that the exported
files could better take advantage of Excel's features.
All text is written with the same text style. Support for additional text styles could be added to highlight certain
data in the file.
The current default style uses fonts that are not on Macs by default so opening the XLSX files in Numbers causes a
pop up that says there are missing fonts. The font could be changed to something that is usually found on Mac and PC.
