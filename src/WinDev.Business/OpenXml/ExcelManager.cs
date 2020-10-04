using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using X15ac = DocumentFormat.OpenXml.Office2013.ExcelAc;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Ds = DocumentFormat.OpenXml.CustomXmlDataProperties;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using Op = DocumentFormat.OpenXml.CustomProperties;

using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WinDev.Common.DataObjects;

namespace WinDev.Business.OpenXml
{
    public class ExcelManager
    {
        StringCollection exceptionMsg = new StringCollection();
        StringBuilder tempMsg = new StringBuilder();
        string completeErrorMessage = string.Empty;


        public void CreateExcelDoc(string fileName, List<RowDataObject> objects)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart1 = document.WorkbookPart.AddNewPart<WorksheetPart>();
                worksheetPart1.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet();
                SheetData sheetDataError = worksheetPart1.Worksheet.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.SheetData());

                Row row = new Row();
                row.Append(
                        ConstructCell("ErrorType", CellValues.String),
                        ConstructCell("ErrorSource", CellValues.String),
                        ConstructCell("ErrorMessage", CellValues.String),
                        ConstructCell("MainErrorMessage", CellValues.String),
                        ConstructCell("FileName", CellValues.String),
                        ConstructCell("Status", CellValues.String),
                        ConstructCell("RawErrorMessage", CellValues.String)
                        );

                sheetDataError.AppendChild(row);

                //foreach (var item in objects.Where(m => m.ProcessStatus == ProcessStatus.Processed))
                //{
                //    row = new Row();

                //    row.Append(
                //        ConstructCell(item.ErrorType, CellValues.Number),
                //        ConstructCell(item.ErrorSource, CellValues.String),
                //        ConstructCell(string.Concat(string.Format("{0}", item.ErrorMessage.TrimEnd())), CellValues.String),
                //        ConstructCell(item.ErrorMessage, CellValues.String),
                //        ConstructCell(item.FileName, CellValues.String),
                //        ConstructCell(item.MessageStatus.ToString(), CellValues.String),
                //        ConstructCell(item.RawErrMessage, CellValues.String));

                //    sheetDataError.AppendChild(row);
                //}

                worksheetPart1.Worksheet.Save();

                // create the worksheet to workbook relation
                document.WorkbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());
                document.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheet()
                {
                    Id = document.WorkbookPart.GetIdOfPart(worksheetPart1),
                    SheetId = 1,
                    Name = "Error List"
                });

                WorksheetPart worksheetPart2 = document.WorkbookPart.AddNewPart<WorksheetPart>();
                worksheetPart2.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet();
                SheetData sheetDataException = worksheetPart2.Worksheet.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.SheetData());

                Row rowException = new Row();
                rowException.Append(
                        ConstructCell("ErrorType", CellValues.String),
                        ConstructCell("ErrorSource", CellValues.String),
                        ConstructCell("ErrorMessage", CellValues.String),
                        ConstructCell("FileName", CellValues.String)
                        );

                sheetDataException.AppendChild(rowException);

                //foreach (var item in objects.Where(m => m.ProcessStatus == ProcessStatus.NotProcessed))
                //{
                //    row = new Row();
                //    row.Append(
                //        ConstructCell(item.ErrorType, CellValues.Number),
                //        ConstructCell(item.ErrorSource, CellValues.String),
                //        ConstructCell(item.ErrorMessage, CellValues.String),
                //        ConstructCell(item.RawErrMessage, CellValues.String),
                //        ConstructCell(item.FileName, CellValues.String));

                //    sheetDataException.AppendChild(row);
                //}

                worksheetPart2.Worksheet.Save();

                document.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheet()
                {
                    Id = document.WorkbookPart.GetIdOfPart(worksheetPart2),
                    SheetId = 2,
                    Name = "Exception List"

                });

                workbookPart.Workbook.Save();
            }
        }

        private IEnumerable<Row> GetRowsFromExcel(string filePath, string sheetName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart workbook = document.WorkbookPart;
                IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

                string sheetId = sheets.Where(p => p.Name == sheetName).FirstOrDefault().Id.Value;
                WorksheetPart worksheetpart = (WorksheetPart)document.WorkbookPart.GetPartById(sheetId);
                SheetData sheetdata = worksheetpart.Worksheet.GetFirstChild<SheetData>();
                return sheetdata.Descendants<Row>();
            }
        }


        public DataSet GetDataFromExcel(string filePath)
        {
            int rowIndex = 1;
            DataSet ds = new DataSet();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart workbook = document.WorkbookPart;
                IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                foreach (Sheet sheet in sheets)
                {
                    WorksheetPart worksheetpart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id.Value);

                    SheetData sheetdata = worksheetpart.Worksheet.GetFirstChild<SheetData>();

                    foreach (Row items in sheetdata.Descendants<Row>())
                    {
                        //DataObject excelMsg = new DataObject();
                        //if (rowIndex != 1)
                        //{
                        //    IEnumerable<Cell> cell = items.Descendants<Cell>();
                        //    excelMsg.ErrorType = cell.Count() >= 1 ? GetCellValue(document, cell.ElementAt(0)) : string.Empty;
                        //    excelMsg.ErrorSource = cell.Count() >= 2 ? GetCellValue(document, cell.ElementAt(1)) : string.Empty;
                        //    excelMsg.ErrorMessage = cell.Count() >= 3 ? GetCellValue(document, cell.ElementAt(2)) : string.Empty;
                        //    excelMsg.MainErrorMessage = cell.Count() >= 4 ? GetCellValue(document, cell.ElementAt(3)) : string.Empty;
                        //    excelMsg.FileName = cell.Count() >= 7 ? GetCellValue(document, cell.ElementAt(6)) : string.Empty;
                        //    excelMsg.MessageStatus = Status.Old;

                        //    excelMsg.RawErrMessage = cell.Count() >= 9 ? GetCellValue(document, cell.ElementAt(8)) : string.Empty;
                        //    errorMsgList.Add(excelMsg);
                        //}
                        rowIndex++;

                    }
                }

            }

            return ds;
        }


        #region "Compare Two different Excel Files"
        private void Export(string mergeFileSavePath, List<RowDataObject> messagesNotMapped)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(mergeFileSavePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart1 = document.WorkbookPart.AddNewPart<WorksheetPart>();
                worksheetPart1.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet();
                SheetData sheetDataError = worksheetPart1.Worksheet.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.SheetData());

                Row row = new Row();
                row.Append(
                        ConstructCell("ErrorCode", CellValues.String),
                                    ConstructCell("Type", CellValues.String),
                                    ConstructCell("ErrorMessage", CellValues.String),
                                    ConstructCell("Issue Description", CellValues.String),
                                    ConstructCell("Issue Resolution", CellValues.String),
                                    ConstructCell("Class", CellValues.String),
                                    ConstructCell("Type", CellValues.String),
                                    ConstructCell("SubType", CellValues.String),
                                    ConstructCell("Workspace", CellValues.String),
                                    ConstructCell("OldErrorMessage", CellValues.String)

                                     );

                sheetDataError.AppendChild(row);

                foreach (var item in messagesNotMapped)
                {

                    row = new Row();
                    //row.Append(
                    //    ConstructCell(item.ErrorCode
                    //    , CellValues.String),
                    //                    ConstructCell(item.ErrorType, CellValues.Number),
                    //                    ConstructCell(item.ErrorMessage, CellValues.String),
                    //                    ConstructCell(item.IssueDescription, CellValues.String),
                    //                    ConstructCell(item.IssueResolution, CellValues.String),
                    //                    ConstructCell(item.Class, CellValues.String),
                    //                    ConstructCell(item.Type, CellValues.String),
                    //                    ConstructCell(item.SubType, CellValues.String),
                    //                    ConstructCell(item.Workspace, CellValues.String),
                    //                    ConstructCell(item.RawErrMessage, CellValues.String));

                    sheetDataError.AppendChild(row);
                }

                worksheetPart1.Worksheet.Save();

                document.WorkbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());

                document.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheet()
                {
                    Id = document.WorkbookPart.GetIdOfPart(worksheetPart1),
                    SheetId = 1,
                    Name = "Sheet1"
                });



                workbookPart.Workbook.Save();
            }
        }

        public DataSet Import(string filePath)
        {
            DataSet ds = new DataSet();
            int rowIndex;

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart workbook = document.WorkbookPart;
                IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                foreach (var sheet in sheets)
                {
                    WorksheetPart worksheetpart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id.Value);
                    SheetData sheetdata = worksheetpart.Worksheet.GetFirstChild<SheetData>();
                    DataTable table = ds.Tables.Add(sheet.Name);
                    rowIndex = 1;

                    foreach (Row items in sheetdata.Descendants<Row>())
                    {
                        DataRow row = table.NewRow();
                        if (rowIndex != 1)
                        {
                            IEnumerable<Cell> cell = items.Descendants<Cell>();
                            row["ErrorCode"] = cell.Count() >= 1 ? GetCellValue(document, cell.ElementAt(0)) : string.Empty;
                            row["ErrorType"] = cell.Count() >= 2 ? GetCellValue(document, cell.ElementAt(1)) : string.Empty;
                            row["ErrorMessage"] = cell.Count() >= 3 ? GetCellValue(document, cell.ElementAt(2)) : string.Empty;
                            row["IssueDescription"] = cell.Count() >= 2 ? GetCellValue(document, cell.ElementAt(3)) : string.Empty;
                            row["IssueResolution"] = cell.Count() >= 5 ? GetCellValue(document, cell.ElementAt(4)) : string.Empty;
                            row["Class"] = cell.Count() >= 6 ? GetCellValue(document, cell.ElementAt(5)) : string.Empty;
                            row["Type"] = cell.Count() >= 7 ? GetCellValue(document, cell.ElementAt(6)) : string.Empty;
                            row["SubType"] = cell.Count() >= 8 ? GetCellValue(document, cell.ElementAt(7)) : string.Empty;
                            row["Workspace"] = cell.Count() >= 9 ? GetCellValue(document, cell.ElementAt(8)) : string.Empty;

                        }
                        rowIndex++;
                    }
                }

            }
            return ds;

        }
        public void ExportDirectoryHierarchy(string directoryPath, string exportFilePath, Func<Files, bool> exclusionFilter)
        {
            FileService service = new FileService();
            List<Files> folderHierarchy = service.GetDirectoryHierarchy(directoryPath, exclusionFilter);            
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(exportFilePath, SpreadsheetDocumentType.Workbook))
            {
                CreateDirectoryHierarchy(document, folderHierarchy);
            }

        }
        #endregion

        #region "OpenXml Document modification methods"
        private Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }
        private Cell ConstructCell(string value, CellValues dataType, uint styleIndex)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                StyleIndex = styleIndex,
                DataType = new EnumValue<CellValues>(dataType)
            };
        }
        private string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            if (cell.CellValue == null)
            {
                return "";
            }
            //int valueint = cell.CellValue.InnerText;
            string value = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }

        #endregion
        #region "Export Directory Hier"        

        // Adds child parts and generates content of the specified part.
        private void CreateDirectoryHierarchy(SpreadsheetDocument document, List<Files> folderHierarchy)
        {

            GenerateExtendedFilePropertiesPart1Content(document);
            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPart1Content(workbookPart1);
            GenerateWorkbookStylesPart1Content(workbookPart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPart1Content(worksheetPart1, folderHierarchy);
            //SharedStringTablePart sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId4");
            //GenerateSharedStringTablePart1Content(sharedStringTablePart1);
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(SpreadsheetDocument document)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Excel";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Worksheets";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "1";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)1U };
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Estimate";

            vTVector2.Append(vTLPSTR2);

            titlesOfParts1.Append(vTVector2);
            Ap.Manager manager1 = new Ap.Manager();
            manager1.Text = "";
            Ap.Company company1 = new Ap.Company();
            company1.Text = "";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinkBase hyperlinkBase1 = new Ap.HyperlinkBase();
            hyperlinkBase1.Text = "";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "16.0300";

            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(manager1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinkBase1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        // Generates content of workbookPart1.
        private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x15 xr xr6 xr10 xr2" } };
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            workbook1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            workbook1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            workbook1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            workbook1.AddNamespaceDeclaration("xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6");
            workbook1.AddNamespaceDeclaration("xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10");
            workbook1.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            FileVersion fileVersion1 = new FileVersion() { ApplicationName = "xl", LastEdited = "7", LowestEdited = "7", BuildVersion = "22527" };
            WorkbookProperties workbookProperties1 = new WorkbookProperties() { DefaultThemeVersion = (UInt32Value)166925U };

            AlternateContent alternateContent1 = new AlternateContent();
            alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "x15" };

            X15ac.AbsolutePath absolutePath1 = new X15ac.AbsolutePath() { Url = "D:\\Projects\\DA\\Simergy\\" };
            absolutePath1.AddNamespaceDeclaration("x15ac", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac");

            alternateContentChoice1.Append(absolutePath1);

            alternateContent1.Append(alternateContentChoice1);

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<xr:revisionPtr revIDLastSave=\"0\" documentId=\"13_ncr:1_{60E8130E-9AC8-48E4-B62A-8EAA12E61EF3}\" xr6:coauthVersionLast=\"45\" xr6:coauthVersionMax=\"45\" xr10:uidLastSave=\"{00000000-0000-0000-0000-000000000000}\" xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\" xmlns:xr6=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision6\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" />");

            BookViews bookViews1 = new BookViews();

            WorkbookView workbookView1 = new WorkbookView() { XWindow = -120, YWindow = -120, WindowWidth = (UInt32Value)20730U, WindowHeight = (UInt32Value)11160U };
            workbookView1.SetAttribute(new OpenXmlAttribute("xr2", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2", "{00000000-000D-0000-FFFF-FFFF00000000}"));

            bookViews1.Append(workbookView1);

            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "Estimate", SheetId = (UInt32Value)8U, Id = "rId1" };

            sheets1.Append(sheet1);
            CalculationProperties calculationProperties1 = new CalculationProperties() { CalculationId = (UInt32Value)191028U };

            WorkbookExtensionList workbookExtensionList1 = new WorkbookExtensionList();

            WorkbookExtension workbookExtension1 = new WorkbookExtension() { Uri = "{140A7094-0E35-4892-8432-C4D2E57EDEB5}" };
            workbookExtension1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.WorkbookProperties workbookProperties2 = new X15.WorkbookProperties() { ChartTrackingReferenceBase = true };

            workbookExtension1.Append(workbookProperties2);

            WorkbookExtension workbookExtension2 = new WorkbookExtension() { Uri = "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" };
            workbookExtension2.AddNamespaceDeclaration("xcalcf", "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures");

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<xcalcf:calcFeatures xmlns:xcalcf=\"http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures\"><xcalcf:feature name=\"microsoft.com:RD\" /><xcalcf:feature name=\"microsoft.com:Single\" /><xcalcf:feature name=\"microsoft.com:FV\" /><xcalcf:feature name=\"microsoft.com:CNMTM\" /></xcalcf:calcFeatures>");

            workbookExtension2.Append(openXmlUnknownElement2);

            workbookExtensionList1.Append(workbookExtension1);
            workbookExtensionList1.Append(workbookExtension2);

            workbook1.Append(fileVersion1);
            workbook1.Append(workbookProperties1);
            workbook1.Append(alternateContent1);
            workbook1.Append(openXmlUnknownElement1);
            workbook1.Append(bookViews1);
            workbook1.Append(sheets1);
            workbook1.Append(calculationProperties1);
            workbook1.Append(workbookExtensionList1);

            workbookPart1.Workbook = workbook1;
        }

        // Generates content of workbookStylesPart1.
        private void GenerateWorkbookStylesPart1Content(WorkbookPart workbookPart1)
        {
            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");

            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac x16r2 xr" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            stylesheet1.AddNamespaceDeclaration("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
            stylesheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)4U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            FontSize fontSize2 = new FontSize() { Val = 10D };
            FontName fontName2 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };

            font2.Append(fontSize2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);

            Font font3 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize3 = new FontSize() { Val = 14D };
            FontName fontName3 = new FontName() { Val = "Tahoma" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };

            font3.Append(bold1);
            font3.Append(fontSize3);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering3);

            Font font4 = new Font();
            FontSize fontSize4 = new FontSize() { Val = 14D };
            FontName fontName4 = new FontName() { Val = "Tahoma" };
            FontFamilyNumbering fontFamilyNumbering4 = new FontFamilyNumbering() { Val = 2 };

            font4.Append(fontSize4);
            font4.Append(fontName4);
            font4.Append(fontFamilyNumbering4);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);

            Fills fills1 = new Fills() { Count = (UInt32Value)11U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            Fill fill3 = new Fill();

            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Indexed = (UInt32Value)22U };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);

            fill3.Append(patternFill3);

            Fill fill4 = new Fill();

            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor2 = new ForegroundColor() { Indexed = (UInt32Value)9U };
            BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill4.Append(foregroundColor2);
            patternFill4.Append(backgroundColor2);

            fill4.Append(patternFill4);

            Fill fill5 = new Fill();

            PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor3 = new ForegroundColor() { Theme = (UInt32Value)4U, Tint = 0.39997558519241921D };
            BackgroundColor backgroundColor3 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill5.Append(foregroundColor3);
            patternFill5.Append(backgroundColor3);

            fill5.Append(patternFill5);

            Fill fill6 = new Fill();

            PatternFill patternFill6 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor4 = new ForegroundColor() { Theme = (UInt32Value)5U, Tint = 0.39997558519241921D };
            BackgroundColor backgroundColor4 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill6.Append(foregroundColor4);
            patternFill6.Append(backgroundColor4);

            fill6.Append(patternFill6);

            Fill fill7 = new Fill();

            PatternFill patternFill7 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor5 = new ForegroundColor() { Theme = (UInt32Value)5U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor5 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill7.Append(foregroundColor5);
            patternFill7.Append(backgroundColor5);

            fill7.Append(patternFill7);

            Fill fill8 = new Fill();

            PatternFill patternFill8 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor6 = new ForegroundColor() { Theme = (UInt32Value)7U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor6 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill8.Append(foregroundColor6);
            patternFill8.Append(backgroundColor6);

            fill8.Append(patternFill8);

            Fill fill9 = new Fill();

            PatternFill patternFill9 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor7 = new ForegroundColor() { Theme = (UInt32Value)9U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor7 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill9.Append(foregroundColor7);
            patternFill9.Append(backgroundColor7);

            fill9.Append(patternFill9);

            Fill fill10 = new Fill();

            PatternFill patternFill10 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor8 = new ForegroundColor() { Theme = (UInt32Value)6U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor8 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill10.Append(foregroundColor8);
            patternFill10.Append(backgroundColor8);

            fill10.Append(patternFill10);

            Fill fill11 = new Fill();

            PatternFill patternFill11 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor9 = new ForegroundColor() { Theme = (UInt32Value)0U, Tint = -4.9989318521683403E-2D };
            BackgroundColor backgroundColor9 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill11.Append(foregroundColor9);
            patternFill11.Append(backgroundColor9);

            fill11.Append(patternFill11);

            fills1.Append(fill1);
            fills1.Append(fill2);
            fills1.Append(fill3);
            fills1.Append(fill4);
            fills1.Append(fill5);
            fills1.Append(fill6);
            fills1.Append(fill7);
            fills1.Append(fill8);
            fills1.Append(fill9);
            fills1.Append(fill10);
            fills1.Append(fill11);

            Borders borders1 = new Borders() { Count = (UInt32Value)2U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color2 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color2);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color3);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color4);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color5 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color5);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)2U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);
            cellStyleFormats1.Append(cellFormat2);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)16U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };

            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment1 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true };

            cellFormat4.Append(alignment1);

            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment2 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true, Indent = (UInt32Value)3U };

            cellFormat5.Append(alignment2);

            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment3 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true, Indent = (UInt32Value)3U };

            cellFormat6.Append(alignment3);

            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)5U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment4 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true, Indent = (UInt32Value)6U };

            cellFormat7.Append(alignment4);

            CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)6U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment5 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true, Indent = (UInt32Value)9U };

            cellFormat8.Append(alignment5);

            CellFormat cellFormat9 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)7U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment6 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true, Indent = (UInt32Value)12U };

            cellFormat9.Append(alignment6);

            CellFormat cellFormat10 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment7 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true, Indent = (UInt32Value)9U };

            cellFormat10.Append(alignment7);

            CellFormat cellFormat11 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment8 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true, Indent = (UInt32Value)12U };

            cellFormat11.Append(alignment8);

            CellFormat cellFormat12 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment9 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true, Indent = (UInt32Value)15U };

            cellFormat12.Append(alignment9);

            CellFormat cellFormat13 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)8U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment10 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true, Indent = (UInt32Value)15U };

            cellFormat13.Append(alignment10);

            CellFormat cellFormat14 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment11 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true, Indent = (UInt32Value)18U };

            cellFormat14.Append(alignment11);

            CellFormat cellFormat15 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)9U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment12 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true, Indent = (UInt32Value)18U };

            cellFormat15.Append(alignment12);

            CellFormat cellFormat16 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment13 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true, Indent = (UInt32Value)21U };

            cellFormat16.Append(alignment13);

            CellFormat cellFormat17 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)10U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment14 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true, Indent = (UInt32Value)21U };

            cellFormat17.Append(alignment14);

            CellFormat cellFormat18 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment15 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true, Indent = (UInt32Value)6U };

            cellFormat18.Append(alignment15);

            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);
            cellFormats1.Append(cellFormat5);
            cellFormats1.Append(cellFormat6);
            cellFormats1.Append(cellFormat7);
            cellFormats1.Append(cellFormat8);
            cellFormats1.Append(cellFormat9);
            cellFormats1.Append(cellFormat10);
            cellFormats1.Append(cellFormat11);
            cellFormats1.Append(cellFormat12);
            cellFormats1.Append(cellFormat13);
            cellFormats1.Append(cellFormat14);
            cellFormats1.Append(cellFormat15);
            cellFormats1.Append(cellFormat16);
            cellFormats1.Append(cellFormat17);
            cellFormats1.Append(cellFormat18);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)2U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            CellStyle cellStyle2 = new CellStyle() { Name = "Normal 2", FormatId = (UInt32Value)1U };
            cellStyle2.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{92B0D24B-3A9A-4A62-A169-CDCA33854AFD}"));

            cellStyles1.Append(cellStyle1);
            cellStyles1.Append(cellStyle2);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles1 = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }

        // Generates content of worksheetPart1.
        private void GenerateWorksheetPart1Content(WorksheetPart worksheetPart1, List<Files> files)
        {

            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac xr xr2 xr3" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            worksheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            worksheet1.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            worksheet1.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
            worksheet1.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{00000000-0001-0000-0600-000000000000}"));
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 69.5703125D, Style = (UInt32Value)1U, CustomWidth = true };

            columns1.Append(column1);
            SheetData sheetData1 = new SheetData();



            foreach (Files file in files)
            {
                Row row = new Row();
                string[] arrObject = file.ToString().Split('|');
                uint styleIndex = (uint)file.column + 1;
                switch (styleIndex)
                {
                    case 2:
                        if (file.hasChildren)
                            styleIndex = 3;

                        break;
                    case 3:
                        styleIndex = !file.hasChildren ? 15u : 4u;
                        break;
                    case 4:
                        styleIndex = !file.hasChildren ? 7u : 5u;
                        break;
                    case 5:
                        styleIndex = !file.hasChildren ? 8u : 6u;
                        break;
                    case 6:
                        styleIndex = !file.hasChildren ? 9u : 10u;
                        break;
                    case 7:
                        styleIndex = !file.hasChildren ? 11u : 12u;
                        break;
                    case 8:
                        styleIndex = !file.hasChildren ? 13u : 14u;
                        break;
                    default:
                        break;
                }
                row.Append(
                  ConstructCell(arrObject.Last(), CellValues.String, styleIndex));
                sheetData1.AppendChild(row);
            }


            #region "Style Referece"            

            //Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = 18D, DyDescent = 0.25D };

            //Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            //CellValue cellValue1 = new CellValue();
            //cellValue1.Text = "0";

            //cell1.Append(cellValue1);

            //row1.Append(cell1);

            //Row row2 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = 18D, DyDescent = 0.25D };

            //Cell cell2 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            //CellValue cellValue2 = new CellValue();
            //cellValue2.Text = "1";

            //cell2.Append(cellValue2);

            //row2.Append(cell2);

            //Row row3 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = 18D, DyDescent = 0.25D };

            //Cell cell3 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            //CellValue cellValue3 = new CellValue();
            //cellValue3.Text = "4";

            //cell3.Append(cellValue3);

            //row3.Append(cell3);

            //Row row4 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = 18D, DyDescent = 0.25D };

            //Cell cell4 = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
            //CellValue cellValue4 = new CellValue();
            //cellValue4.Text = "5";

            //cell4.Append(cellValue4);

            //row4.Append(cell4);

            //Row row5 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = 18D, DyDescent = 0.25D };

            //Cell cell5 = new Cell() { CellReference = "A5", StyleIndex = (UInt32Value)15U, DataType = CellValues.SharedString };
            //CellValue cellValue5 = new CellValue();
            //cellValue5.Text = "6";

            //cell5.Append(cellValue5);

            //row5.Append(cell5);

            //Row row6 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = 18D, DyDescent = 0.25D };

            //Cell cell6 = new Cell() { CellReference = "A6", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            //CellValue cellValue6 = new CellValue();
            //cellValue6.Text = "7";

            //cell6.Append(cellValue6);

            //row6.Append(cell6);

            //Row row7 = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = 18D, DyDescent = 0.25D };

            //Cell cell7 = new Cell() { CellReference = "A7", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            //CellValue cellValue7 = new CellValue();
            //cellValue7.Text = "8";

            //cell7.Append(cellValue7);

            //row7.Append(cell7);

            //Row row8 = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = 18D, DyDescent = 0.25D };

            //Cell cell8 = new Cell() { CellReference = "A8", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            //CellValue cellValue8 = new CellValue();
            //cellValue8.Text = "2";

            //cell8.Append(cellValue8);

            //row8.Append(cell8);

            //Row row9 = new Row() { RowIndex = (UInt32Value)9U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = 18D, DyDescent = 0.25D };

            //Cell cell9 = new Cell() { CellReference = "A9", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            //CellValue cellValue9 = new CellValue();
            //cellValue9.Text = "3";

            //cell9.Append(cellValue9);

            //row9.Append(cell9);

            //Row row10 = new Row() { RowIndex = (UInt32Value)10U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = 18D, DyDescent = 0.25D };

            //Cell cell10 = new Cell() { CellReference = "A10", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
            //CellValue cellValue10 = new CellValue();
            //cellValue10.Text = "2";

            //cell10.Append(cellValue10);

            //row10.Append(cell10);

            //Row row11 = new Row() { RowIndex = (UInt32Value)11U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = 18D, DyDescent = 0.25D };

            //Cell cell11 = new Cell() { CellReference = "A11", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            //CellValue cellValue11 = new CellValue();
            //cellValue11.Text = "3";

            //cell11.Append(cellValue11);

            //row11.Append(cell11);

            //Row row12 = new Row() { RowIndex = (UInt32Value)12U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = 18D, DyDescent = 0.25D };

            //Cell cell12 = new Cell() { CellReference = "A12", StyleIndex = (UInt32Value)10U, DataType = CellValues.SharedString };
            //CellValue cellValue12 = new CellValue();
            //cellValue12.Text = "2";

            //cell12.Append(cellValue12);

            //row12.Append(cell12);

            //Row row13 = new Row() { RowIndex = (UInt32Value)13U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = 18D, DyDescent = 0.25D };

            //Cell cell13 = new Cell() { CellReference = "A13", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            //CellValue cellValue13 = new CellValue();
            //cellValue13.Text = "3";

            //cell13.Append(cellValue13);

            //row13.Append(cell13);

            //Row row14 = new Row() { RowIndex = (UInt32Value)14U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = 18D, DyDescent = 0.25D };

            //Cell cell14 = new Cell() { CellReference = "A14", StyleIndex = (UInt32Value)12U, DataType = CellValues.SharedString };
            //CellValue cellValue14 = new CellValue();
            //cellValue14.Text = "2";

            //cell14.Append(cellValue14);

            //row14.Append(cell14);

            //Row row15 = new Row() { RowIndex = (UInt32Value)15U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = 18D, DyDescent = 0.25D };

            //Cell cell15 = new Cell() { CellReference = "A15", StyleIndex = (UInt32Value)13U, DataType = CellValues.SharedString };
            //CellValue cellValue15 = new CellValue();
            //cellValue15.Text = "3";

            //cell15.Append(cellValue15);

            //row15.Append(cell15);

            //Row row16 = new Row() { RowIndex = (UInt32Value)16U, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = 18D, DyDescent = 0.25D };

            //Cell cell16 = new Cell() { CellReference = "A16", StyleIndex = (UInt32Value)14U, DataType = CellValues.SharedString };
            //CellValue cellValue16 = new CellValue();
            //cellValue16.Text = "2";

            //cell16.Append(cellValue16);

            //row16.Append(cell16);

            //sheetData1.Append(row1);
            //sheetData1.Append(row2);
            //sheetData1.Append(row3);
            //sheetData1.Append(row4);
            //sheetData1.Append(row5);
            //sheetData1.Append(row6);
            //sheetData1.Append(row7);
            //sheetData1.Append(row8);
            //sheetData1.Append(row9);
            //sheetData1.Append(row10);
            //sheetData1.Append(row11);
            //sheetData1.Append(row12);
            //sheetData1.Append(row13);
            //sheetData1.Append(row14);
            //sheetData1.Append(row15);
            //sheetData1.Append(row16);



            #endregion

            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup1 = new PageSetup() { PaperSize = (UInt32Value)9U, Orientation = OrientationValues.Portrait, VerticalDpi = (UInt32Value)0U, Id = "rId1" };
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(pageMargins1);
            worksheet1.Append(pageSetup1);

            worksheetPart1.Worksheet = worksheet1;
        }

        // Generates content of sharedStringTablePart1.
        private void GenerateSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1)
        {
            SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)6U, UniqueCount = (UInt32Value)6U };

            SharedStringItem sharedStringItem1 = new SharedStringItem();
            Text text1 = new Text();
            text1.Text = "3. Development & Testing";

            sharedStringItem1.Append(text1);

            SharedStringItem sharedStringItem2 = new SharedStringItem();
            Text text2 = new Text();
            text2.Text = "Development Milestone -1";

            sharedStringItem2.Append(text2);

            SharedStringItem sharedStringItem3 = new SharedStringItem();
            Text text3 = new Text();
            text3.Text = "Access Management";

            sharedStringItem3.Append(text3);

            SharedStringItem sharedStringItem4 = new SharedStringItem();
            Text text4 = new Text();
            text4.Text = "Create Roles";

            sharedStringItem4.Append(text4);

            SharedStringItem sharedStringItem5 = new SharedStringItem();
            Text text5 = new Text();
            text5.Text = "2. Requirement Prototyping";

            sharedStringItem5.Append(text5);

            SharedStringItem sharedStringItem6 = new SharedStringItem();
            Text text6 = new Text();
            text6.Text = "Screen Prototypes";

            sharedStringItem6.Append(text6);

            sharedStringTable1.Append(sharedStringItem1);
            sharedStringTable1.Append(sharedStringItem2);
            sharedStringTable1.Append(sharedStringItem3);
            sharedStringTable1.Append(sharedStringItem4);
            sharedStringTable1.Append(sharedStringItem5);
            sharedStringTable1.Append(sharedStringItem6);

            sharedStringTablePart1.SharedStringTable = sharedStringTable1;
        }




        #endregion
    }
}
