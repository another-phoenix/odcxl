using OfficeOpenXml;
using OutSystems.ExternalLibraries.SDK;
using System.Net;

namespace ODCXL
{
    public class ODCXL : IODCXL
    {
        public ODCXL()
        {
        }

        void IODCXL.Cell_CalculateByIndex(object worksheet, int row, int column)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Cell_CalculateByName(object worksheet, string cellName)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Cell_Read(object worksheet, string cellName, int cellRow, int cellColumn, out string cellValue, bool readText)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Cell_ReadByIndex(object worksheet, int row, int column, bool readText, out string cellValue)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Cell_ReadByName(object worksheet, string cellName, bool readText, out string cellValue)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Cell_SetFormulaByIndex(object worksheet, int row, int column, string formula)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Cell_SetFormulaByName(object worksheet, string cellName, string formula)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Cell_WriteByIndex(object worksheet, int row, int column, string cellValue, string cellType)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Cell_WriteByName(object worksheet, string cellName, string cellValue, string cellType)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Cell_WriteImageByIndex(object worksheet, int row, int column, string ssImageName, byte[] ssImage)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Cell_WriteImageByName(object worksheet, string cellName, string ssImageName, byte[] ssImage)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Column_Delete(object worksheet, int startColumnNumber, int numberOfColumns)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Column_Hide_Show(object worksheet, int column, bool hidden)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Column_Insert(object worksheet, int insertAt, int numberOfColumns, int copyStylesFrom)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Column_SetWidth(object worksheet, int columnNumber, decimal ssDesiredWidth)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Comment_Add(object worksheet, int rowNumber, int columnNumber, string text, string author, bool autofit, bool isRichText)
        {
            throw new NotImplementedException();
        }

        void IODCXL.ConditionalFormatting_DeleteAllRules(object worksheet)
        {
            throw new NotImplementedException();
        }

        void IODCXL.ConditionalFormatting_DeleteRule(object worksheet, int ssRuleToDeleteIndex)
        {
            throw new NotImplementedException();
        }

        void IODCXL.ContainInRange(object worksheet, string range, string value, string parameter1, out bool found, out int rowIndex, out int columnIndex)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Image_Insert(object worksheet, byte[] ssImageFile, string ssImageType, string ssImageName, int rowNumber, int columnNumber, string cellName, int ssImageWidth, int ssImageHeight, int ssMarginTop, int ssMarginLeft)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Row_Delete(object worksheet, int startRowNumber, int numberOfRows)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Row_Hide_Show(object worksheet, int rowIndex, bool hidden)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Row_Insert(object worksheet, int insertAt, int nrRows, int copyStyleFromRow)
        {
            throw new NotImplementedException();
        }

        void IODCXL.Row_SetHeight(object worksheet, int rowNumber, decimal ssDesiredHeight)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkBook_AddCopyWorkSheet(object workBook, string worksheetName, object worksheetToCopy, out object worksheet)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkBook_AddName(object workBook, string worksheetName, out object worksheet)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkBook_AddSheet(object workBook, string worksheetName, object worksheet, int indexWhereToAdd)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkBook_Calculate(object workBook)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkBook_ChangeSheetIndex(object workBook, int currentIndex, int newIndex)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkBook_Close(object workBook)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkBook_GetBinaryData(object workBook, out byte[] ssBinaryData)
        {
            throw new NotImplementedException();
        }

        public byte[] WorkBook_Open(string fileName, byte[] binaryData)
        {
            if (binaryData.LongLength <= 0 && string.IsNullOrEmpty(fileName))
            {
                throw new Exception("You need to specify at least one of FileName or Binary_Data");
            }

            ExcelPackage p = new ExcelPackage();
            if (fileName.ToLower().StartsWith("http:") || fileName.ToLower().StartsWith("https:"))
            {
                System.Net.HttpWebRequest request = (HttpWebRequest)WebRequest.Create(fileName);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                p.Load(response.GetResponseStream());
            }
            else if (!string.IsNullOrEmpty(fileName))
            {
                p.Load(System.IO.File.Open(fileName, System.IO.FileMode.OpenOrCreate));
            }
            else if (binaryData.LongLength > 0)
            {
                Stream s = new MemoryStream(binaryData);
                p.Load(s);
            }
            else
            {
                throw new FileNotFoundException("Could not open a file with the given information. Please verify your filename/binary data and try again.");
            }
            return p.GetAsByteArray();
        }

        public byte[] WorkBook_Open_BinaryData(byte[] ssBinaryData)
        {
            return WorkBook_Open("", ssBinaryData);
        }

        void IODCXL.WorkBook_Protect(object workBook, string password, bool ssLockStructure, bool ssLockWindows, bool ssLockRevision)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkSheet_AddName(object worksheet, string name, object dataSet, int rowStart, int columnStart)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkSheet_AutofitColumns(object worksheet)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkSheet_Calculate(object worksheet)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkSheet_Delete(object workBook, int indexToDelete, string nameToDelete)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkSheet_DeleteByIndex(object workBook, int indexToDelete)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkSheet_DeleteByName(object workBook, string nameToDelete)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkSheet_GetName(object worksheet, out string worksheetName)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkSheet_Hide_Show(object worksheet, int hidden)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkSheet_Rename(object worksheet, string name)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkSheet_Select(object workBook, int worksheetIndex, string worksheetName, out object worksheet)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkSheet_SelectByIndex(object workBook, int worksheetNumber, out object worksheet)
        {
            throw new NotImplementedException();
        }

        void IODCXL.WorkSheet_SelectByName(object workBook, string worksheetName, out object worksheet)
        {
            throw new NotImplementedException();
        }
    }
}
