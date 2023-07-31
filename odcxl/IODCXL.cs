using OutSystems.ExternalLibraries.SDK;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ODCXL
{
    [OSInterface]
    public interface IODCXL
    {


        /// <summary>
        /// Opens an existing workbook for editing by either specifying a name or the binary data.
        /// </summary>
        /// <param name="fileName">Location of the file that you want to open. Set to empty string &quot;&quot; when using binary data</param>
        /// <param name="binaryData">Binary data of the file that you want to open. Set to nullbinary() if using FileName</param>
        /// <param name="workBook">The workbook that you want to work with.</param>
        byte[] WorkBook_Open(string fileName, byte[] binaryData);

        /// <summary>
        /// Select a worksheet by its index or by its name
        /// </summary>
        /// <param name="workBook">The workbook wherein the worksheet exists</param>
        /// <param name="worksheetIndex">The index of the worksheet to find. Indexes start at 1</param>
        /// <param name="worksheetName">The name of the worksheet to find</param>
        /// <param name="worksheet">This is the worksheet object that you have been looking for,</param>
        void WorkSheet_Select(object workBook, int worksheetIndex, string worksheetName, out object worksheet);

        /// <summary>
        /// Creates a new excel workbook, optionally specifying the name of the fiirst sheet.
        /// </summary>
        /// <param name="numberOfSheets">The number of sheets to add. Sheet names will be auto generated, i.e. Sheet1, Sheet2.</param>
        /// <param name="firstSheetName">Specify the name of the initial sheet in the workbook. Default = &quot;Sheet1&quot;</param>
        /// <param name="sheetNames">List of new sheets to add, with at least a name specified. The index, if specified, will be used to add sheets in that order.
        /// FirstSheetName and NrSheets are ignored if SheetNames is populated</param>
        /// <param name="workBook">The newly created workbook</param>
        //void WorkBook_Create(int numberOfSheets, string firstSheetName, RLNewSheetRecordList sheetNames, out object workBook);

        /// <summary>
        /// Get the in-memory binary data of the specified workbook
        /// </summary>
        /// <param name="workBook">The workbook you want the binary data for</param>
        /// <param name="ssBinaryData">The binary data of the file</param>
        void WorkBook_GetBinaryData(object workBook, out byte[] ssBinaryData);

        /// <summary>
        /// Rename a worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="name">The new name for the spreadsheet</param>
        void WorkSheet_Rename(object worksheet, string name);

        /// <summary>
        /// Closes the excel workbook
        /// </summary>
        /// <param name="workBook"></param>
        void WorkBook_Close(object workBook);

        /// <summary>
        /// Hides / Shows a Column passed by index
        /// </summary>
        /// <param name="worksheet">The worksheet you want to work with.</param>
        /// <param name="column">The index of the column within the worksheet that you want to hide/show.</param>
        /// <param name="hidden">A Boolean value, set to True to hide the column, and to False to show the column.</param>
        void Column_Hide_Show(object worksheet, int column, bool hidden);

        /// <summary>
        /// Reads the value of a cell.
        /// </summary>
        /// <param name="worksheet">WorkSheet on which the cell resides</param>
        /// <param name="cellName">Name of the cell to read from, i.e. A4. Required if CellRow and CellNumber set to 0.</param>
        /// <param name="cellRow">Row number of the cell to read from. Required if CellName not set.</param>
        /// <param name="cellColumn">Column number of the cell to read from. Required if CellName not set.</param>
        /// <param name="cellValue">The value in the cell, as text.</param>
        /// <param name="readText">If true always reads the cell value as text</param>
        void Cell_Read(object worksheet, string cellName, int cellRow, int cellColumn, out string cellValue, bool readText);

        /// <summary>
        /// Set protection on an Excel WorkSheet
        /// </summary>
        /// <param name="worksheet">WorkSheet to protect</param>
        /// <param name="password">DEPRECATED
        /// Can be used for backwards compatibility with Excel_Package
        /// 
        /// Paword to protect the worksheet with.
        /// </param>
        /// <param name="protectionOptions">Options to set when protecting the worksheet</param>
        //void WorkSheet_Protect(object worksheet, string password, RCProtectionRecord protectionOptions);

        /// <summary>
        /// Write a converted value to a cell.
        /// </summary>
        /// <param name="worksheet">WorkSheet on which the cell resides </param>
        /// <param name="cellName">Name of the cell to write to, i.e. A4. Required if CellRow and CellColumn not set</param>
        /// <param name="cellRow">Row number of the cell to write to. Required if CellName not set.</param>
        /// <param name="cellColumn">Column number of the cell to write to. Required if CellName not set.</param>
        /// <param name="cellValue">The value to write to the cell</param>
        /// <param name="cellType">Type can be:
        /// text (default),
        /// datetime,
        /// integer,
        /// decimal,
        /// boolean,
        /// formula</param>
        /// <param name="cellFormat">CellFormat for the target cell</param>
        //void Cell_Write(object worksheet, string cellName, int cellRow, int cellColumn, string cellValue, string cellType, RCCellFormatRecord cellFormat);

        /// <summary>
        /// Write a dataset to a range of cells.
        /// Accepts format for the target cells
        /// </summary>
        /// <param name="worksheet">WorkSheet to write to</param>
        /// <param name="rowStart">Start row (integer)</param>
        /// <param name="columnStart">Start column (integer)</param>
        /// <param name="dataSet">Data to write</param>
        /// <param name="cellFormat">CellFormat for the target cells</param>
        /// <param name="exportHeaders">True to include headers in export file. Default value = False</param>
        //void Cell_WriteRange(object worksheet, int rowStart, int columnStart, object dataSet, RCCellFormatRecord cellFormat, bool exportHeaders);

        /// <summary>
        /// Get the name of the given worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="worksheetName"></param>
        void WorkSheet_GetName(object worksheet, out string worksheetName);

        /// <summary>
        /// Get all properties of the workbook
        /// </summary>
        /// <param name="workBook">The workbook</param>
        /// <param name="properties"></param>
        //void WorkBook_GetProperties(object workBook, out RCWorkbookRecord properties);

        /// <summary>
        /// Get the properties of the given worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="properties"></param>
        //void WorkSheet_GetProperties(object worksheet, out RCWorkSheetRecord properties);

        /// <summary>
        /// Add a worksheet to an existing workbook, optionally at the index specified. Specifying only a name will create a blank sheet. Specifying  a name with binary data, will add the sheet from the existing binary data, and then rename to the newly provided name
        /// </summary>
        /// <param name="workBook">The workbook that you want to add the sheet to</param>
        /// <param name="worksheetName">The name of the worksheet you want to add. If binary data is nullbinary(), an empty sheet will be added</param>
        /// <param name="worksheet">The worksheet object that you want to add. Set to nullbinary() if adding a new sheet by name</param>
        /// <param name="indexWhereToAdd">The index where to add the new sheet. Default will be highest sheet index plus 1</param>
        void WorkBook_AddSheet(object workBook, string worksheetName, object worksheet, int indexWhereToAdd);

        /// <summary>
        /// Delete a worksheet in a workbook by specifying either the index, or the name of the worksheet.
        /// </summary>
        /// <param name="workBook">The workbook from which you want to delete the worksheet</param>
        /// <param name="indexToDelete">The index of the worksheet to delete. Set to 0 if using the worksheet name to delete</param>
        /// <param name="nameToDelete">The name of the worksheet to delete. Set to empty string &quot;&quot; if using the index to delete.</param>
        void WorkSheet_Delete(object workBook, int indexToDelete, string nameToDelete);

        /// <summary>
        /// Create a chart
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="chartType">Receives the chart type in text, possible types:
        /// Area3D
        /// AreaStacked3D
        /// AreaStacked1003D
        /// BarClustered3D
        /// BarStacked3D
        /// BarStacked1003D
        /// Column3D
        /// ColumnClustered3D
        /// ColumnStacked3D
        /// ColumnStacked1003D
        /// Line3D
        /// Pie3D
        /// PieExploded3D
        /// Area
        /// AreaStacked
        /// AreaStacked100
        /// BarClustered
        /// BarOfPie
        /// BarStacked
        /// BarStacked100
        /// Bubble
        /// Bubble3DEffect
        /// ColumnClustered
        /// ColumnStacked
        /// ColumnStacked100
        /// ConeBarClustered
        /// ConeBarStacked
        /// ConeBarStacked100
        /// ConeCol
        /// ConeColClustered
        /// ConeColStacked
        /// ConeColStacked100
        /// CylinderBarClustered
        /// CylinderBarStacked
        /// CylinderBarStacked100
        /// CylinderCol
        /// CylinderColClustered
        /// CylinderColStacked
        /// CylinderColStacked100
        /// Doughnut
        /// DoughnutExploded
        /// Line
        /// LineMarkers
        /// LineMarkersStacked
        /// LineMarkersStacked100
        /// LineStacked
        /// LineStacked100
        /// Pie
        /// PieExploded
        /// PieOfPie
        /// PyramidBarClustered
        /// PyramidBarStacked
        /// PyramidBarStacked100
        /// PyramidCol
        /// PyramidColClustered
        /// PyramidColStacked
        /// PyramidColStacked100
        /// Radar
        /// RadarFilled
        /// RadarMarkers
        /// StockHLC
        /// StockOHLC
        /// StockVHLC
        /// StockVOHLC
        /// Surface
        /// SurfaceTopView
        /// SurfaceTopViewWireframe
        /// SurfaceWireframe
        /// XYScatter
        /// XYScatterLines
        /// XYScatterLinesNoMarkers
        /// XYScatterSmooth
        /// XYScatterSmoothNoMarkers=73</param>
        /// <param name="chartName"></param>
        /// <param name="dataSeries_List">List Of DataSeries</param>
        /// <param name="height">Expressed in pixels</param>
        /// <param name="width">Expressed in pixels</param>
        /// <param name="rowPos">Row position to place the upper left corner graph</param>
        /// <param name="colPos">Column position to place the upper left corner graph</param>
        //void Chart_Create(object worksheet, string chartType, string chartName, RLDataSeriesRecordList dataSeries_List, int height, int width, int rowPos, int colPos);

        /// <summary>
        /// Inserts a new row into the spreadsheet.  Existing rows below the position are shifted down.  All formula are updated to take account of the new row.
        /// </summary>
        /// <param name="worksheet">The worksheet to insert the row(s) into</param>
        /// <param name="insertAt">The position of the new row
        /// </param>
        /// <param name="nrRows">Number of rows to insert</param>
        /// <param name="copyStyleFromRow">Copy Styles from this row. Applied to all inserted rows. 0 will not copy any styles</param>
        void Row_Insert(object worksheet, int insertAt, int nrRows, int copyStyleFromRow);

        /// <summary>
        /// Change the index of a worksheet in the document
        /// </summary>
        /// <param name="workBook">The workbook in which the change is to be made.</param>
        /// <param name="currentIndex">The current index(position) of the sheet in question</param>
        /// <param name="newIndex">The new index for the sheet</param>
        void WorkBook_ChangeSheetIndex(object workBook, int currentIndex, int newIndex);

        /// <summary>
        /// Apply a specified cell format to the range specified for the given worksheet
        /// </summary>
        /// <param name="worksheet">WorkSheet object where formatting is to be applied</param>
        /// <param name="cellFormat">CellFormat to apply</param>
        /// <param name="range">Range that CellFormat is to be applied to</param>
        //void CellFormat_ApplyToRange(object worksheet, RCCellFormatRecord cellFormat, RCRangeRecord range);

        /// <summary>
        /// Find all cells that contain the specified value in the given worksheet
        /// </summary>
        /// <param name="worksheet">The worksheet in which to search</param>
        /// <param name="valueToFind">The value to search for</param>
        /// <param name="listOfCells">List of cells (ranges) where the value has been found</param>
        //void Cells_FindByValue(object worksheet, string valueToFind, out RLRangeRecordList listOfCells);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="range"></param>
        /// <param name="value"></param>
        /// <param name="parameter1"></param>
        /// <param name="found"></param>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        void ContainInRange(object worksheet, string range, string value, string parameter1, out bool found, out int rowIndex, out int columnIndex);

        /// <summary>
        /// Calculate all formulae for the entire workbook provided.
        /// </summary>
        /// <param name="workBook">The workbook to work with</param>
        void WorkBook_Calculate(object workBook);

        /// <summary>
        /// Calculate all formulae on the provided worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet to work with</param>
        void WorkSheet_Calculate(object worksheet);

        /// <summary>
        /// Hides / Shows Row passed by index
        /// </summary>
        /// <param name="worksheet">WorkSheet to work with</param>
        /// <param name="rowIndex">Index of the Row to show/hide</param>
        /// <param name="hidden">A Boolean value, set to True to hide the row and to False to show the row</param>
        void Row_Hide_Show(object worksheet, int rowIndex, bool hidden);

        /// <summary>
        /// Hide / Show a worksheet
        /// </summary>
        /// <param name="worksheet">The worksheet to work with</param>
        /// <param name="hidden">Visible = 0 - The worksheet is visible
        /// Hidden = 1 - The worksheet is hidden but can be shown by the user via the user interface
        /// VeryHidden = 2 - The worksheet is hidden and cannot be shown by the user via the user interface</param>
        void WorkSheet_Hide_Show(object worksheet, int hidden);

        /// <summary>
        /// Add a rule for conditionally formatting a range of cells.
        /// </summary>
        /// <param name="worksheet">The worksheet to work with</param>
        /// <param name="conditionalFormatRecord">The conditional formatting to apply to the Address Range</param>
        //void ConditionalFormatting_AddRule(object worksheet, RCConditionalFormatItemRecord conditionalFormatRecord);

        /// <summary>
        /// Get a list of all the conditional formatting rules in a worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet to work with</param>
        /// <param name="listOfRule">List of conditional formatting rules</param>
        //void ConditionalFormatting_GetAllRules(object worksheet, out RLConditionalFormatItemRecordList listOfRule);

        /// <summary>
        /// Merge cells in the range provided
        /// </summary>
        /// <param name="worksheet">The worksheet to work with</param>
        /// <param name="rangeToMerge">The range of the cells to merge</param>
        //void Cell_Merge(object worksheet, RCRangeRecord rangeToMerge);

        /// <summary>
        /// Un-Merge cells in the range provided
        /// </summary>
        /// <param name="worksheet">The worksheet to work with</param>
        /// <param name="rangeToUnmerge">The range of cell to un-merge</param>
        //void Cell_UnMerge(object worksheet, RCRangeRecord rangeToUnmerge);

        /// <summary>
        /// Delete row(s) from a worksheet
        /// </summary>
        /// <param name="worksheet">The worksheet to work with</param>
        /// <param name="startRowNumber">Row number where to start deleting rows.</param>
        /// <param name="numberOfRows">The number of rows to delete. Default = 1.</param>
        void Row_Delete(object worksheet, int startRowNumber, int numberOfRows);

        /// <summary>
        /// Delete column(s) from a worksheet
        /// </summary>
        /// <param name="worksheet">The worksheet to work with</param>
        /// <param name="startColumnNumber">Column number where to start deleting columns.</param>
        /// <param name="numberOfColumns">The number of rows to delete. Default = 1.</param>
        void Column_Delete(object worksheet, int startColumnNumber, int numberOfColumns);

        /// <summary>
        /// Delete comment(s) in a specified range
        /// </summary>
        /// <param name="worksheet">The worksheet to work with.</param>
        /// <param name="range">Range to delete comments from.</param>
        //void Comment_Delete(object worksheet, RCRangeRecord range);

        /// <summary>
        /// Add a comment to a cell
        /// </summary>
        /// <param name="worksheet">The worksheet to work with.</param>
        /// <param name="rowNumber">The row number of the cell to add the comment to.</param>
        /// <param name="columnNumber">The column number of the cell to add the comment to.</param>
        /// <param name="text">The comment.</param>
        /// <param name="author">The author of the comment.</param>
        /// <param name="autofit">True to autofit the comment window to the comment text</param>
        /// <param name="isRichText">Is the comment rich text</param>
        void Comment_Add(object worksheet, int rowNumber, int columnNumber, string text, string author, bool autofit, bool isRichText);

        /// <summary>
        /// Inserts a new column into the spreadsheet.  Existing columns to the right of the insert index will be shifted right.  All formula are updated to take account of the new column.
        /// </summary>
        /// <param name="worksheet">The worksheet to work with.</param>
        /// <param name="insertAt">Column number where to insert new column.</param>
        /// <param name="numberOfColumns">The number of columns to insert.</param>
        /// <param name="copyStylesFrom">Copy Styles from this column. Applied to all inserted columns. 0 (default) will not copy any styles</param>
        void Column_Insert(object worksheet, int insertAt, int numberOfColumns, int copyStylesFrom);

        /// <summary>
        /// Delete a specified Conditional Formatting rule on a worksheet
        /// </summary>
        /// <param name="worksheet">The worksheet to work with.</param>
        /// <param name="ssRuleToDeleteIndex">The index of the rule to be deleted.</param>
        void ConditionalFormatting_DeleteRule(object worksheet, int ssRuleToDeleteIndex);

        /// <summary>
        /// Delete ALL Conditional Formatting rules for a worksheet
        /// </summary>
        /// <param name="worksheet">The worksheet to work with.</param>
        void ConditionalFormatting_DeleteAllRules(object worksheet);

        /// <summary>
        /// Insert an image into a WorkSheet
        /// </summary>
        /// <param name="worksheet">The worksheet to work with</param>
        /// <param name="ssImageFile">Binary data of the image to be inserted</param>
        /// <param name="ssImageType">File type. BMP, PNG, JPG</param>
        /// <param name="ssImageName">Name reference for the image in the WorkSheet</param>
        /// <param name="rowNumber">Row index where to insert image. Ignored if CellName is specified</param>
        /// <param name="columnNumber">Column index where to insert image. Ignored if CellName is specified</param>
        /// <param name="cellName">Cell Name where to insert image</param>
        /// <param name="ssImageWidth">The width of the image in pixels</param>
        /// <param name="ssImageHeight">The height of the image in pixels</param>
        /// <param name="ssMarginTop"> Offset in pixels	</param>
        /// <param name="ssMarginLeft"> Offset in pixels</param>
        void Image_Insert(object worksheet, byte[] ssImageFile, string ssImageType, string ssImageName, int rowNumber, int columnNumber, string cellName, int ssImageWidth, int ssImageHeight, int ssMarginTop, int ssMarginLeft);

        /// <summary>
        /// Apply the column autofit action to the specified range of cells specified in the given worksheet
        /// </summary>
        /// <param name="worksheet">The worksheet to work with</param>
        void WorkSheet_AutofitColumns(object worksheet);

        /// <summary>
        /// Add the automatic filter option of Excel to the specified range of cells.
        /// </summary>
        /// <param name="worksheet">The worksheet to work with.</param>
        /// <param name="rangeToFilter">The range where to add the filter. If not supplied, the dimension of the worksheet will be used.</param>
        //void WorkSheet_AddAutoFilter(object worksheet, RCRangeRecord rangeToFilter);

        /// <summary>
        /// Set protection on the workbook level
        /// </summary>
        /// <param name="workBook">The workbook to work with</param>
        /// <param name="password">The paword to set for the workbook. This does not encrypt the workbook.</param>
        /// <param name="ssLockStructure">Locks the structure,which prevents users from adding or deleting worksheets or from displaying hidden worksheets.</param>
        /// <param name="ssLockWindows">Locks the position of the workbook window.</param>
        /// <param name="ssLockRevision">Lock the workbook for revision</param>
        void WorkBook_Protect(object workBook, string password, bool ssLockStructure, bool ssLockWindows, bool ssLockRevision);

        /// <summary>
        /// Calculates the formula of a cell, defined by its index.
        /// </summary>
        /// <param name="worksheet">WorkSheet on which the cell resides</param>
        /// <param name="row">Row Number</param>
        /// <param name="column">Column Number</param>
        void Cell_CalculateByIndex(object worksheet, int row, int column);

        /// <summary>
        /// Calculates the formula of a cell, defined by its name.
        /// </summary>
        /// <param name="worksheet">WorkSheet on which the cell resides</param>
        /// <param name="cellName">Cell-name (eg A4)</param>
        void Cell_CalculateByName(object worksheet, string cellName);

        /// <summary>
        /// Apply format to a range of cells.
        /// </summary>
        /// <param name="worksheet">WorkSheet to write to</param>
        /// <param name="rowStart">Start row (integer)</param>
        /// <param name="columnStart">Start column (integer)</param>
        /// <param name="rowEnd">End row (integer)</param>
        /// <param name="columnEnd">End column (integer)</param>
        /// <param name="cellFormat">CellFormat for the target cells</param>
        //void Cell_FormatRange(object worksheet, int rowStart, int columnStart, int rowEnd, int columnEnd, RCCellFormatRecord cellFormat);

        /// <summary>
        /// Reads the value of a cell, defined by its index.
        /// </summary>
        /// <param name="worksheet">WorkSheet on which the cell resides</param>
        /// <param name="row">row number</param>
        /// <param name="column">column number</param>
        /// <param name="readText">If true always reads the cell value as text</param>
        /// <param name="cellValue">text-value</param>
        void Cell_ReadByIndex(object worksheet, int row, int column, bool readText, out string cellValue);

        /// <summary>
        /// Reads the value of a cell, defined by its name.
        /// </summary>
        /// <param name="worksheet">WorkSheet on which the cell resides</param>
        /// <param name="cellName">Cell-name (eg A4)</param>
        /// <param name="readText">If true always reads the cell value as text</param>
        /// <param name="cellValue">text-value</param>
        void Cell_ReadByName(object worksheet, string cellName, bool readText, out string cellValue);

        /// <summary>
        /// Write a formula to a cell, defined by its index.
        /// </summary>
        /// <param name="worksheet">WorkSheet on which the cell resides</param>
        /// <param name="row">rownumber</param>
        /// <param name="column">columnnumber</param>
        /// <param name="formula">Formula</param>
        void Cell_SetFormulaByIndex(object worksheet, int row, int column, string formula);

        /// <summary>
        /// Write a formula to a cell, defined by its name.
        /// </summary>
        /// <param name="worksheet">WorkSheet on which the cell resides</param>
        /// <param name="cellName">Cell-name (eg A4)</param>
        /// <param name="formula">Formula</param>
        void Cell_SetFormulaByName(object worksheet, string cellName, string formula);

        /// <summary>
        /// Adds a copy of a worksheet
        /// </summary>
        /// <param name="workBook">The workbook in which the worksheet is to be copied
        /// </param>
        /// <param name="worksheetName">The name of the spreadsheet to create</param>
        /// <param name="worksheetToCopy">The worksheet to be copied</param>
        /// <param name="worksheet">The copied worksheet</param>
        void WorkBook_AddCopyWorkSheet(object workBook, string worksheetName, object worksheetToCopy, out object worksheet);

        /// <summary>
        /// Get all images in a worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="ssImages"></param>
        //void WorkSheet_GetImages(object worksheet, out RLImageRecordList ssImages);

        /// <summary>
        /// Select a worksheet by its index
        /// </summary>
        /// <param name="workBook">The worksheet to work with</param>
        /// <param name="worksheetNumber">The index of the spreadsheet to select, starting at 1</param>
        /// <param name="worksheet">The selected worksheet</param>
        void WorkSheet_SelectByIndex(object workBook, int worksheetNumber, out object worksheet);

        /// <summary>
        /// Select a worksheet to work on by its name
        /// </summary>
        /// <param name="workBook">The workbook to work with</param>
        /// <param name="worksheetName">The name of the spreadsheet to select</param>
        /// <param name="worksheet">The selected worksheet</param>
        void WorkSheet_SelectByName(object workBook, string worksheetName, out object worksheet);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="nameToDelete"></param>
        void WorkSheet_DeleteByName(object workBook, string nameToDelete);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="indexToDelete"></param>
        void WorkSheet_DeleteByIndex(object workBook, int indexToDelete);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="worksheet">The worksheet you want to work with.</param>
        /// <param name="chartType">Receives the chart type in text, possible types:
        /// Area3D
        /// AreaStacked3D
        /// AreaStacked1003D
        /// BarClustered3D
        /// BarStacked3D
        /// BarStacked1003D
        /// Column3D
        /// ColumnClustered3D
        /// ColumnStacked3D
        /// ColumnStacked1003D
        /// Line3D
        /// Pie3D
        /// PieExploded3D
        /// Area
        /// AreaStacked
        /// AreaStacked100
        /// BarClustered
        /// BarOfPie
        /// BarStacked
        /// BarStacked100
        /// Bubble
        /// Bubble3DEffect
        /// ColumnClustered
        /// ColumnStacked
        /// ColumnStacked100
        /// ConeBarClustered
        /// ConeBarStacked
        /// ConeBarStacked100
        /// ConeCol
        /// ConeColClustered
        /// ConeColStacked
        /// ConeColStacked100
        /// CylinderBarClustered
        /// CylinderBarStacked
        /// CylinderBarStacked100
        /// CylinderCol
        /// CylinderColClustered
        /// CylinderColStacked
        /// CylinderColStacked100
        /// Doughnut
        /// DoughnutExploded
        /// Line
        /// LineMarkers
        /// LineMarkersStacked
        /// LineMarkersStacked100
        /// LineStacked
        /// LineStacked100
        /// Pie
        /// PieExploded
        /// PieOfPie
        /// PyramidBarClustered
        /// PyramidBarStacked
        /// PyramidBarStacked100
        /// PyramidCol
        /// PyramidColClustered
        /// PyramidColStacked
        /// PyramidColStacked100
        /// Radar
        /// RadarFilled
        /// RadarMarkers
        /// StockHLC
        /// StockOHLC
        /// StockVHLC
        /// StockVOHLC
        /// Surface
        /// SurfaceTopView
        /// SurfaceTopViewWireframe
        /// SurfaceWireframe
        /// XYScatter
        /// XYScatterLines
        /// XYScatterLinesNoMarkers
        /// XYScatterSmooth
        /// XYScatterSmoothNoMarkers=73</param>
        /// <param name="chartName"></param>
        /// <param name="dataSeries_List">List Of DataSeries</param>
        /// <param name="height">Expressed in pixels</param>
        /// <param name="width">Expressed in pixels</param>
        /// <param name="rowPos">Row position to place the upper left corner graph</param>
        /// <param name="colPos">Column position to place the upper left corner graph</param>
        //void WorkSheet_Chart_Create(object worksheet, string chartType, string chartName, RLDataSeriesRecordList dataSeries_List, int height, int width, int rowPos, int colPos);

        /// <summary>
        /// Create a defined &quot;Name&quot; (a word or string of characters in Excel that represents a cell, range of cells, formula, or constant value) in excel, starting in the RowStart / ColumnStart cell.
        /// </summary>
        /// <param name="worksheet">WorkSheet to write to</param>
        /// <param name="name">&quot;Name&quot;</param>
        /// <param name="dataSet">Values to assigned the name</param>
        /// <param name="rowStart">Start row number</param>
        /// <param name="columnStart">Start column number</param>
        void WorkSheet_AddName(object worksheet, string name, object dataSet, int rowStart, int columnStart);

        /// <summary>
        /// Opens an existing workbook for editing and keeps it in memory
        /// </summary>
        /// <param name="binaryData"></param>
        /// <param name="workBook"></param>
        byte[] WorkBook_Open_BinaryData(byte[] binaryData);

        /// <summary>
        /// Set the pixel width of a column on a specific worksheet
        /// </summary>
        /// <param name="worksheet">The worksheet to work with</param>
        /// <param name="columnNumber">The column number, starting at 1</param>
        /// <param name="ssDesiredWidth">The pixel width you desire for the column.</param>
        void Column_SetWidth(object worksheet, int columnNumber, decimal ssDesiredWidth);

        /// <summary>
        /// Set the pixel height for a specific row in a worksheet
        /// </summary>
        /// <param name="worksheet">The worksheet to work with
        /// </param>
        /// <param name="rowNumber">The number of the row to set the height for</param>
        /// <param name="ssDesiredHeight">The desired pixel height for the row</param>
        void Row_SetHeight(object worksheet, int rowNumber, decimal ssDesiredHeight);

        /// <summary>
        /// Write a converted value to a cell, defined by its index.
        /// </summary>
        /// <param name="worksheet">WorkSheet on which the cell resides</param>
        /// <param name="row">Row Number</param>
        /// <param name="column">Column Number</param>
        /// <param name="cellValue">Text Value</param>
        /// <param name="cellType">Type can by text (default), datetime, integer, decimal, boolean</param>
        void Cell_WriteByIndex(object worksheet, int row, int column, string cellValue, string cellType);

        /// <summary>
        /// Write a converted value to a cell, defined by its index.
        /// Accepts format for the target cell
        /// </summary>
        /// <param name="worksheet">WorkSheet on which the cell resides</param>
        /// <param name="row">Row Number</param>
        /// <param name="column">Column Number</param>
        /// <param name="cellValue">Text Value</param>
        /// <param name="cellType">Type can by text (default), datetime, integer, decimal, boolean</param>
        /// <param name="cellFormat">CellFormat for the target cell</param>
        //void Cell_WriteByIndexWithFormat(object worksheet, int row, int column, string cellValue, string cellType, RCCellFormatRecord cellFormat);

        /// <summary>
        /// Write a converted value to a cell, defined by its name.
        /// </summary>
        /// <param name="worksheet">WorkSheet in which the cell resides</param>
        /// <param name="cellName">Cell-name (eg A4)</param>
        /// <param name="cellValue">Value to write</param>
        /// <param name="cellType">Type can by text (default), datetime, integer, decimal, boolean</param>
        void Cell_WriteByName(object worksheet, string cellName, string cellValue, string cellType);

        /// <summary>
        /// Write a converted value to a cell, defined by its name.
        /// Accepts format for the target cell
        /// </summary>
        /// <param name="worksheet">WorkSheet in which the cell resides</param>
        /// <param name="cellName">Cell-name (eg A4)</param>
        /// <param name="cellValue">Value to write</param>
        /// <param name="cellType">Type can by text (default), datetime, integer, decimal, boolean</param>
        /// <param name="cellFormat">CellFormat for the target cell</param>
       // void Cell_WriteByNameWithFormat(object worksheet, string cellName, string cellValue, string cellType, RCCellFormatRecord cellFormat);

        /// <summary>
        /// Write a dataset to a range of column cells
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="row"></param>
        /// <param name="columnStart"></param>
        /// <param name="valueList"></param>
        /// <param name="cellType"></param>
        //void Cell_WriteColumnRange(object worksheet, int row, int columnStart, RLValueRecordList valueList, string cellType);

        /// <summary>
        /// Write a dataset to a range of column cells
        /// Accepts format for the target cells
        /// </summary>
        /// <param name="worksheet">WorkSheet to write to</param>
        /// <param name="row">rownumber</param>
        /// <param name="columnStart">Start column (integer)</param>
        /// <param name="valueList">Values to write to columns</param>
        /// <param name="cellType">Type can by text (default), datetime, integer, decimal, boolean</param>
        /// <param name="cellFormat">CellFormat for the target cells</param>
        //void Cell_WriteColumnRangeWithFormat(object worksheet, int row, int columnStart, RLValueRecordList valueList, string cellType, RCCellFormatRecord cellFormat);

        /// <summary>
        /// Write a image on a cell, defined by its index.
        /// </summary>
        /// <param name="worksheet">WorkSheet on which the cell resides</param>
        /// <param name="row">row number</param>
        /// <param name="column">column number</param>
        /// <param name="ssImageName">The image name</param>
        /// <param name="ssImage">The image to write.</param>
        void Cell_WriteImageByIndex(object worksheet, int row, int column, string ssImageName, byte[] ssImage);

        /// <summary>
        /// Write a image on a cell, defined by its name.
        /// </summary>
        /// <param name="worksheet">WorkSheet on which the cell resides</param>
        /// <param name="cellName">Cell-name (eg A4)</param>
        /// <param name="ssImageName">The image name</param>
        /// <param name="ssImage">The image to write.</param>
        void Cell_WriteImageByName(object worksheet, string cellName, string ssImageName, byte[] ssImage);

        /// <summary>
        /// Write a dataset to a range of cells.
        /// Accepts format for the target cells
        /// </summary>
        /// <param name="worksheet">WorkSheet to write to</param>
        /// <param name="rowStart">Start row (integer)</param>
        /// <param name="columnStart">Start column (integer)</param>
        /// <param name="dataSet">Data to write</param>
        /// <param name="cellFormat">CellFormat for the target cells</param>
        //void Cell_WriteRangeWithFormat(object worksheet, int rowStart, int columnStart, object dataSet, RCCellFormatRecord cellFormat);

        /// <summary>
        /// Add a worksheet to work on by its name
        /// </summary>
        /// <param name="workBook">Workbook where the sheet is to be added</param>
        /// <param name="worksheetName">The name of the spreadsheet to create</param>
        /// <param name="worksheet">The newly added worksheet</param>
        void WorkBook_AddName(object workBook, string worksheetName, out object worksheet);

  
}
}
