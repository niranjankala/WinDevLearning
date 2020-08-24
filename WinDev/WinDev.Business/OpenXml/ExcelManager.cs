using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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

        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
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


        private List<string> SeparateErrorMessage(string message)
        {
            message = Regex.Replace(message.Trim(), @"\*\*+", @"*"); //kbErr.ErrorMessage.Trim().Replace("**", "*");
            message = Regex.Replace(message, "  +", " ");
            List<string> regExMsgParts = message.Split('*').ToList();
            regExMsgParts.RemoveAll(str => string.IsNullOrWhiteSpace(str) || str.Trim().Length == 1);
            return regExMsgParts;
        }

        private bool MatchPatternWithSimulationError(string errorMessage, List<string> regExMsgParts)
        {
            errorMessage = Regex.Replace(errorMessage, "  +", " ");
            return regExMsgParts.All(msgPart => msgPart.Trim().Length > 1 && errorMessage.Contains(msgPart.Trim()));
        }
        public bool MatchPatternWithSimulationError(string kbMessage, string errorMessage)
        {
            List<string> regExMsgParts = SeparateErrorMessage(kbMessage); // xml msg parts
            errorMessage = Regex.Replace(errorMessage, "  +", " ");
            //int cc = regExMsgParts.Count(x => x.Trim().Length > 1 && errorMessage.Contains(x.Trim()));
            //if(cc >= regExMsgParts.Count-1 && regExMsgParts.Count >= cc && regExMsgParts.Distinct().Count() == regExMsgParts.Count())
            //{
            //    return true;
            //}
            return regExMsgParts.All(msgPart => msgPart.Trim().Length > 1 && errorMessage.Contains(msgPart.Trim()));
        }

        private string CreateRegExFromKBError(string kbErrorInfoMessage)
        {
            string regExString = kbErrorInfoMessage;
            if (!string.IsNullOrWhiteSpace(kbErrorInfoMessage))
            {

                StringBuilder sbResult = new StringBuilder(Regex.Escape(kbErrorInfoMessage));
                sbResult.Replace("\\*]", "\\*\\]");
                //sbResult.Replace("\\*", "[^$]*");
                //sbResult.Replace("\\*", "[^ ]*"); //Not working if space found b/w variable section
                sbResult.Replace("\\*", ".*?");

                regExString = sbResult.ToString();
                regExString = Regex.Replace(regExString, @"\\ \\ +", "\\s*?");
                regExString = Regex.Replace(regExString, string.Format("({0})+", Regex.Escape(".*?")), ".*?");
                regExString = Regex.Replace(regExString, string.Format("({0})+", Regex.Escape(".*?\\s*?")), ".*?");
                regExString = Regex.Replace(regExString, string.Format("({0})+", Regex.Escape("\\s*?")), "\\s*?");
                if (regExString.EndsWith(".*?"))
                {
                    regExString = regExString.Substring(0, regExString.Length - 1);
                }
                //regExString = regExString.StartsWith(".*?") ? regExString : regExString.Insert(0, ".*?");
            }
            return regExString;
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
        public void ExportDirectoryHierarchy(string directoryPath, string exportFilePath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(exportFilePath, SpreadsheetDocumentType.Workbook))
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

                //foreach (var item in messagesNotMapped)
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
        #endregion
    }
}
