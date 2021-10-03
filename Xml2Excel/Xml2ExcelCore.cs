using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Xml.Serialization;

namespace Xml2Excel
{
    public class Xml2ExcelCore
    {
        public Xml2ExcelCore()
        {

        }

        public bool Generate(string xml,string outputPath)
        {
            try
            {
                Workbook workbook=GetWorkbookFromXml(xml);
                var xlWorkbook = GenerateWorkbook(workbook);
                xlWorkbook.SaveAs(outputPath);
                return true;
            }
            catch (Exception e)
            {
                throw new Xml2ExcelException(e.Message,e);
            }
        }

        public MemoryStream Generate(string xml)
        {
            try
            {
                Workbook workbook = GetWorkbookFromXml(xml);
                var xlWorkbook = GenerateWorkbook(workbook);
                using (MemoryStream stream = new MemoryStream())
                {
                    xlWorkbook.SaveAs(stream);
                    return stream;
                }
            }
            catch (Exception e)
            {
                throw new Xml2ExcelException(e.Message,e);
            }
        }

        private IXLWorkbook GenerateWorkbook(Workbook workbook)
        {
            try
            {
                var xlWorkbook = new XLWorkbook();
                if (!string.IsNullOrEmpty(workbook.author))
                    xlWorkbook.Properties.Author = workbook.author;
                if (!string.IsNullOrEmpty(workbook.title))
                    xlWorkbook.Properties.Title = workbook.title;
                if (!string.IsNullOrEmpty(workbook.subject))
                    xlWorkbook.Properties.Subject = workbook.subject;
                if (!string.IsNullOrEmpty(workbook.category))
                    xlWorkbook.Properties.Category = workbook.category;
                if (!string.IsNullOrEmpty(workbook.keywords))
                    xlWorkbook.Properties.Keywords = workbook.keywords;
                if (!string.IsNullOrEmpty(workbook.comments))
                    xlWorkbook.Properties.Comments = workbook.comments;
                if (!string.IsNullOrEmpty(workbook.status))
                    xlWorkbook.Properties.Status = workbook.status;
                if (!string.IsNullOrEmpty(workbook.company))
                    xlWorkbook.Properties.Company = workbook.company;
                if (!string.IsNullOrEmpty(workbook.manager))
                    xlWorkbook.Properties.Manager = workbook.manager;

                foreach (var worksheetItem in workbook.Worksheets)
                {
                    GenerateWorksheet(worksheetItem, xlWorkbook);
                }
                xlWorkbook.CalculateMode = XLCalculateMode.Auto;
                return xlWorkbook;
            }
            catch (Exception e)
            {

                throw new Xml2ExcelException(e.Message,e);
            }
        }

        private void GenerateWorksheet(Worksheet worksheet, XLWorkbook xLWorkbook)
        {
            try
            {
                IXLWorksheet xLWorksheet = xLWorkbook.Worksheets.Add(string.IsNullOrEmpty(worksheet.name) ? string.Format("Tab {0}", xLWorkbook.Worksheets.Count + 1) : worksheet.name);
                if (!string.IsNullOrEmpty(worksheet.tabColor))
                    xLWorksheet.SetTabColor(XLColor.FromHtml(worksheet.tabColor));
                if (worksheet.rowHeight > 0)
                    xLWorksheet.RowHeight = worksheet.rowHeight;

                if (!string.IsNullOrEmpty(worksheet.password))
                {
                    xLWorksheet.Protect(worksheet.password);
                }

                foreach (var item in worksheet.Ranges)
                {
                    GenerateRange(item, xLWorksheet);
                }

                foreach (var item in worksheet.Rows)
                {
                    GenerateRow(item, xLWorksheet);
                }

                foreach (var item in worksheet.Columns)
                {
                    GenerateColumn(item, xLWorksheet);
                }

                foreach (var item in worksheet.Cells)
                {
                    GenerateCell(item, xLWorksheet);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                //throw new Xml2ExcelException(e.Message, e);
            }
            

            
        }

        private async void GenerateCell(Cell cell,IXLWorksheet xLWorksheet)
        {
            try
            {
                //if (cell.row == 0 || cell.column == 0) throw new Xml2ExcelException("Row or column values must be greater than 0");
                IXLCell xlCell = xLWorksheet.Cell(cell.row, cell.column);
                if (!string.IsNullOrEmpty(cell.link))
                    xlCell.Hyperlink = new XLHyperlink(cell.link);

                if (!string.IsNullOrEmpty(cell.extLink))
                    xlCell.Hyperlink.ExternalAddress = new Uri(cell.extLink);

                if (!string.IsNullOrEmpty(cell.numberFormat))
                    xlCell.Style.NumberFormat.Format = cell.numberFormat;

                xlCell.Style.NumberFormat.SetNumberFormatId(cell.formatId);

                if (!string.IsNullOrEmpty(cell.image))
                {
                    using (WebClient webClient = new WebClient())
                    {
                        byte[] data = webClient.DownloadData(cell.image);
                        using (Stream stream = new MemoryStream(data))
                        {
                            var image = xLWorksheet.AddPicture(stream)//(cell.image)
                            .MoveTo(xlCell);
                            if (cell.imageScale != 0)
                                image.Scale(cell.imageScale); // optional: resize picture
                        }

                    }
                }
                if (!string.IsNullOrEmpty(cell.style))
                    GenerateStyles(xlCell, cell.style);

                if (!string.IsNullOrEmpty(cell.formula))
                    xlCell.SetFormulaA1(cell.formula);
                else xlCell.Value = cell.value;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                //throw new Xml2ExcelException(e.Message, e);
            }
            
        }

        private Workbook GetWorkbookFromXml(string xml)
        {
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Workbook));
                using (TextReader reader = new StringReader(xml))
                {
                    return (Workbook)serializer.Deserialize(reader);
                }
            }
            catch (Exception e)
            {

                throw new Xml2ExcelException(e.Message, e);
            }
            
        }

        private void GenerateStyles(IXLCell xLCell, string styles)
        {
            try
            {
                string[] styleList = styles.Split(';');
                var borderStyle = new Dictionary<string, int>()
                {
                    { "dotted",(int)XLBorderStyleValues.Dotted},
                    { "solid",(int)XLBorderStyleValues.Thick},
                    { "dashed",(int)XLBorderStyleValues.Dashed},
                    { "none",(int)XLBorderStyleValues.None},
                };
                    var backgroundColorPattern = new Dictionary<string, int>()
                {
                    { "solid",(int)XLFillPatternValues.Solid},
                    { "darktrellis",(int)XLFillPatternValues.DarkTrellis},
                    { "lighttrellis",(int)XLFillPatternValues.LightTrellis},
                    { "darkhorizontal",(int)XLFillPatternValues.DarkHorizontal},
                    { "lighthorizontal",(int)XLFillPatternValues.LightHorizontal},
                    { "darkvertical",(int)XLFillPatternValues.DarkVertical},
                    { "lightvertical",(int)XLFillPatternValues.LightVertical},
                    { "darkdown",(int)XLFillPatternValues.DarkDown},
                    { "lightdown",(int)XLFillPatternValues.LightDown},
                    { "darkup",(int)XLFillPatternValues.DarkUp},
                    { "lightup",(int)XLFillPatternValues.LightUp},
                    { "lightgray",(int)XLFillPatternValues.LightGray},
                    { "darkgray",(int)XLFillPatternValues.DarkGray},
                    { "darkgrid",(int)XLFillPatternValues.DarkGrid},
                    { "lightgrid",(int)XLFillPatternValues.LightGrid},
                    { "none",(int)XLFillPatternValues.None},
                };
                foreach (var item in styleList)
                {
                    string[] keyValue = item.Split(':');
                    if (keyValue.Length != 2) continue;
                    string key = keyValue[0].Trim();
                    string value = keyValue[1].Trim();
                    switch (key)
                    {
                        case "text-align":
                            {
                                if (value == "left")
                                    xLCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                                else if (value == "right")
                                    xLCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                                else if (value == "center")
                                    xLCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            }
                            break;
                        case "border-color":
                            {
                                xLCell.Style.Border.SetLeftBorderColor(XLColor.FromHtml(value));
                                xLCell.Style.Border.LeftBorder = (xLCell.Style.Border.LeftBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLCell.Style.Border.LeftBorder;

                                xLCell.Style.Border.SetRightBorderColor(XLColor.FromHtml(value));
                                xLCell.Style.Border.RightBorder = (xLCell.Style.Border.RightBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLCell.Style.Border.RightBorder;

                                xLCell.Style.Border.SetTopBorderColor(XLColor.FromHtml(value));
                                xLCell.Style.Border.TopBorder = (xLCell.Style.Border.TopBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLCell.Style.Border.TopBorder;

                                xLCell.Style.Border.SetBottomBorderColor(XLColor.FromHtml(value));
                                xLCell.Style.Border.BottomBorder = (xLCell.Style.Border.BottomBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLCell.Style.Border.BottomBorder;
                            }
                            break;
                        case "border-left-color":
                            {
                                xLCell.Style.Border.SetLeftBorderColor(XLColor.FromHtml(value));
                                xLCell.Style.Border.LeftBorder = (xLCell.Style.Border.LeftBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLCell.Style.Border.LeftBorder;

                            }
                            break;
                        case "border-right-color":
                            {
                                xLCell.Style.Border.SetRightBorderColor(XLColor.FromHtml(value));
                                xLCell.Style.Border.RightBorder = (xLCell.Style.Border.RightBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLCell.Style.Border.RightBorder;
                            }
                            break;
                        case "border-top-color":
                            {
                                xLCell.Style.Border.SetTopBorderColor(XLColor.FromHtml(value));
                                xLCell.Style.Border.TopBorder = (xLCell.Style.Border.TopBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLCell.Style.Border.TopBorder;
                            }
                            break;
                        case "border-bottom-color":
                            {
                                xLCell.Style.Border.SetBottomBorderColor(XLColor.FromHtml(value));
                                xLCell.Style.Border.BottomBorder = (xLCell.Style.Border.BottomBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLCell.Style.Border.BottomBorder;
                            }
                            break;
                        case "border-style":
                            {
                                if (!borderStyle.ContainsKey(value)) continue;
                                borderStyle.TryGetValue(value, out int borderStyleValue);
                                xLCell.Style.Border.TopBorder = (XLBorderStyleValues)borderStyleValue;
                                xLCell.Style.Border.RightBorder = (XLBorderStyleValues)borderStyleValue;
                                xLCell.Style.Border.BottomBorder = (XLBorderStyleValues)borderStyleValue;
                                xLCell.Style.Border.LeftBorder = (XLBorderStyleValues)borderStyleValue;
                            }
                            break;
                        case "border-top-style":
                            {
                                if (!borderStyle.ContainsKey(value)) continue;
                                borderStyle.TryGetValue(value, out int borderStyleValue);
                                xLCell.Style.Border.TopBorder = (XLBorderStyleValues)borderStyleValue;
                            }
                            break;
                        case "border-right-style":
                            {
                                if (!borderStyle.ContainsKey(value)) continue;
                                borderStyle.TryGetValue(value, out int borderStyleValue);
                                xLCell.Style.Border.RightBorder = (XLBorderStyleValues)borderStyleValue;
                            }
                            break;
                        case "border-bottom-style":
                            {
                                if (!borderStyle.ContainsKey(value)) continue;
                                borderStyle.TryGetValue(value, out int borderStyleValue);
                                xLCell.Style.Border.BottomBorder = (XLBorderStyleValues)borderStyleValue;
                            }
                            break;
                        case "border-left-style":
                            {
                                if (!borderStyle.ContainsKey(value)) continue;
                                borderStyle.TryGetValue(value, out int borderStyleValue);
                                xLCell.Style.Border.LeftBorder = (XLBorderStyleValues)borderStyleValue;
                            }
                            break;
                        case "background-color":
                            {
                                xLCell.Style.Fill.PatternType = (xLCell.Style.Fill.PatternType == XLFillPatternValues.None) ? XLFillPatternValues.Solid : xLCell.Style.Fill.PatternType;
                                xLCell.Style.Fill.BackgroundColor = XLColor.FromHtml(value);
                            }
                            break;
                        case "background-color-pattern":
                            {
                                if (!backgroundColorPattern.ContainsKey(value)) continue;
                                backgroundColorPattern.TryGetValue(value, out int backgroundColorPatternValue);
                                xLCell.Style.Fill.PatternType = (XLFillPatternValues)backgroundColorPatternValue;
                            }
                            break;
                        case "color":
                            {
                                xLCell.Style.Font.FontColor = XLColor.FromHtml(value);
                            }
                            break;
                        case "font-style":
                            {
                                switch (value)
                                {
                                    case "italic": xLCell.Style.Font.Italic = true; break;
                                    case "shadow": xLCell.Style.Font.Shadow = true; break;
                                }
                            }
                            break;
                        case "font-weight":
                            {
                                switch (value)
                                {
                                    case "bold": xLCell.Style.Font.SetBold(true); break;
                                    case "normal": xLCell.Style.Font.Bold = false; break;
                                }
                            }
                            break;
                        case "text-decoration":
                            {
                                switch (value)
                                {
                                    case "none": xLCell.Style.Font.SetUnderline(XLFontUnderlineValues.None); break;
                                    case "underline": xLCell.Style.Font.SetUnderline(XLFontUnderlineValues.Single); break;
                                    case "underline-double": xLCell.Style.Font.SetUnderline(XLFontUnderlineValues.Double); break;
                                    case "strikethrough": xLCell.Style.Font.SetStrikethrough(true); break;
                                }
                            }
                            break;
                        case "font-size":
                            {
                                xLCell.Style.Font.FontSize = double.Parse(value);
                            }
                            break;
                        case "font-family":
                            {
                                xLCell.Style.Font.FontName = value;
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                //throw new Xml2ExcelException(e.Message, e);
            }
            
        }

        private void GenerateRange(Range range, IXLWorksheet xLWorksheet)
        {
            try
            {
                string[] cell1 = range.cell1.Split(',');
                if (cell1.Length != 2) throw new Exception("The field \"cell1\" of the range must have the format \"row, column\"");
                int cell1Row = int.Parse(cell1[0]);
                int cell1Column = int.Parse(cell1[1]);
                string[] cell2 = range.cell2.Split(',');
                if (cell2.Length != 2) throw new Exception("The field \"cell2\" of the range must have the format \"row, column\"");
                int cell2Row = int.Parse(cell2[0]);
                int cell2Column = int.Parse(cell2[1]);
                IXLRange xLRange = xLWorksheet.Range(cell1Row, cell1Column, cell2Row, cell2Column);
                if (range.merge)
                    xLRange.Merge();
                if (range.clear)
                    xLRange.Clear();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                //throw new Xml2ExcelException(e.Message, e);
            }
            
        }


        private void GenerateRow(Row row, IXLWorksheet xLWorksheet)
        {
            try
            {
                IXLRow xLRow = xLWorksheet.Row(row.number);
                if (row.height > 0)
                    xLRow.Height = row.height;
                if (row.adjustToContents)
                    xLRow.AdjustToContents();
                if (!string.IsNullOrEmpty(row.style))
                {
                    GenerateRowStyles(xLRow, row.style);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                //throw new Xml2ExcelException(e.Message, e);
            }
            
        }

        private void GenerateRowStyles(IXLRow xLRow, string styles)
        {
            try
            {
                string[] styleList = styles.Split(';');
                var borderStyle = new Dictionary<string, int>()
                {
                    { "dotted",(int)XLBorderStyleValues.Dotted},
                    { "solid",(int)XLBorderStyleValues.Thick},
                    { "dashed",(int)XLBorderStyleValues.Dashed},
                    { "none",(int)XLBorderStyleValues.None},
                };
                var backgroundColorPattern = new Dictionary<string, int>()
                {
                    { "solid",(int)XLFillPatternValues.Solid},
                    { "darktrellis",(int)XLFillPatternValues.DarkTrellis},
                    { "lighttrellis",(int)XLFillPatternValues.LightTrellis},
                    { "darkhorizontal",(int)XLFillPatternValues.DarkHorizontal},
                    { "lighthorizontal",(int)XLFillPatternValues.LightHorizontal},
                    { "darkvertical",(int)XLFillPatternValues.DarkVertical},
                    { "lightvertical",(int)XLFillPatternValues.LightVertical},
                    { "darkdown",(int)XLFillPatternValues.DarkDown},
                    { "lightdown",(int)XLFillPatternValues.LightDown},
                    { "darkup",(int)XLFillPatternValues.DarkUp},
                    { "lightup",(int)XLFillPatternValues.LightUp},
                    { "lightgray",(int)XLFillPatternValues.LightGray},
                    { "darkgray",(int)XLFillPatternValues.DarkGray},
                    { "darkgrid",(int)XLFillPatternValues.DarkGrid},
                    { "lightgrid",(int)XLFillPatternValues.LightGrid},
                    { "none",(int)XLFillPatternValues.None},
                };
                foreach (var item in styleList)
                {
                    string[] keyValue = item.Split(':');
                    if (keyValue.Length != 2) continue;
                    string key = keyValue[0].Trim();
                    string value = keyValue[1].Trim();
                    switch (key)
                    {
                        case "text-align":
                        {
                            if (value == "left")
                                xLRow.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                            else if (value == "right")
                                xLRow.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                            else if(value == "center")
                                xLRow.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        }break;
                        case "border-color":
                        {
                            xLRow.Style.Border.SetLeftBorderColor(XLColor.FromHtml(value));
                            xLRow.Style.Border.LeftBorder= (xLRow.Style.Border.LeftBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLRow.Style.Border.LeftBorder;

                            xLRow.Style.Border.SetRightBorderColor(XLColor.FromHtml(value));
                            xLRow.Style.Border.RightBorder = (xLRow.Style.Border.RightBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLRow.Style.Border.RightBorder;

                            xLRow.Style.Border.SetTopBorderColor(XLColor.FromHtml(value));
                            xLRow.Style.Border.TopBorder = (xLRow.Style.Border.TopBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLRow.Style.Border.TopBorder;

                            xLRow.Style.Border.SetBottomBorderColor(XLColor.FromHtml(value));
                            xLRow.Style.Border.BottomBorder = (xLRow.Style.Border.BottomBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLRow.Style.Border.BottomBorder;
                        }
                            break;
                        case "border-left-color":
                        {
                            xLRow.Style.Border.SetLeftBorderColor(XLColor.FromHtml(value));
                            xLRow.Style.Border.LeftBorder = (xLRow.Style.Border.LeftBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLRow.Style.Border.LeftBorder;

                        }break;
                        case "border-right-color":
                        {
                            xLRow.Style.Border.SetRightBorderColor(XLColor.FromHtml(value));
                            xLRow.Style.Border.RightBorder = (xLRow.Style.Border.RightBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLRow.Style.Border.RightBorder;
                        }break;
                        case "border-top-color":
                        {
                            xLRow.Style.Border.SetTopBorderColor(XLColor.FromHtml(value));
                            xLRow.Style.Border.TopBorder = (xLRow.Style.Border.TopBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLRow.Style.Border.TopBorder;
                        }break;
                        case "border-bottom-color":
                        {
                            xLRow.Style.Border.SetBottomBorderColor(XLColor.FromHtml(value));
                            xLRow.Style.Border.BottomBorder =(xLRow.Style.Border.BottomBorder== XLBorderStyleValues.None)? XLBorderStyleValues.Thick: xLRow.Style.Border.BottomBorder;
                        }break;
                        case "border-style":
                        {
                            if (!borderStyle.ContainsKey(value)) continue;
                            borderStyle.TryGetValue(value, out int borderStyleValue);
                            xLRow.Style.Border.TopBorder = (XLBorderStyleValues)borderStyleValue;
                            xLRow.Style.Border.RightBorder = (XLBorderStyleValues)borderStyleValue;
                            xLRow.Style.Border.BottomBorder = (XLBorderStyleValues)borderStyleValue;
                            xLRow.Style.Border.LeftBorder = (XLBorderStyleValues)borderStyleValue;
                        }break;
                        case "border-top-style":
                        {
                            if (!borderStyle.ContainsKey(value)) continue;
                            borderStyle.TryGetValue(value, out int borderStyleValue);
                            xLRow.Style.Border.TopBorder = (XLBorderStyleValues)borderStyleValue;
                        }break;
                        case "border-right-style":
                        {
                            if (!borderStyle.ContainsKey(value)) continue;
                            borderStyle.TryGetValue(value, out int borderStyleValue);
                            xLRow.Style.Border.RightBorder = (XLBorderStyleValues)borderStyleValue;
                        }break;
                        case "border-bottom-style":
                        {
                            if (!borderStyle.ContainsKey(value)) continue;
                            borderStyle.TryGetValue(value, out int borderStyleValue);
                            xLRow.Style.Border.BottomBorder = (XLBorderStyleValues)borderStyleValue;
                        }break;
                        case "border-left-style":
                        {
                            if (!borderStyle.ContainsKey(value)) continue;
                            borderStyle.TryGetValue(value, out int borderStyleValue);
                            xLRow.Style.Border.LeftBorder = (XLBorderStyleValues)borderStyleValue;
                        }break;
                        case "background-color":
                        {
                            xLRow.Style.Fill.PatternType =(xLRow.Style.Fill.PatternType== XLFillPatternValues.None)? XLFillPatternValues.Solid: xLRow.Style.Fill.PatternType;
                            xLRow.Style.Fill.BackgroundColor = XLColor.FromHtml(value);
                        }break;
                        case "background-color-pattern":
                        {
                            if (!backgroundColorPattern.ContainsKey(value)) continue;
                            backgroundColorPattern.TryGetValue(value, out int backgroundColorPatternValue);
                            xLRow.Style.Fill.PatternType =(XLFillPatternValues)backgroundColorPatternValue;
                        }break;
                        case "color":
                        {
                            xLRow.Style.Font.FontColor = XLColor.FromHtml(value);
                        }break;
                        case "font-style":
                        {
                            switch (value)
                            {
                                case "italic": xLRow.Style.Font.Italic = true;break;
                                case "shadow": xLRow.Style.Font.Shadow=true;break;
                            }
                        }break;
                        case "font-weight":
                        {
                            switch (value)
                            {
                                case "bold": xLRow.Style.Font.SetBold(true); break;
                                case "normal": xLRow.Style.Font.Bold=false;break;
                            }
                        }break;
                        case "text-decoration":
                        {
                            switch (value)
                            {
                                case "none": xLRow.Style.Font.SetUnderline(XLFontUnderlineValues.None); break;
                                case "underline": xLRow.Style.Font.SetUnderline(XLFontUnderlineValues.Single); break;
                                case "underline-double": xLRow.Style.Font.SetUnderline(XLFontUnderlineValues.Double); break;
                                case "strikethrough": xLRow.Style.Font.SetStrikethrough(true); break;
                            }
                        }break;
                        case "font-size":
                        {
                            xLRow.Style.Font.FontSize=double.Parse(value);
                        }break;
                        case "font-family":
                        {
                            xLRow.Style.Font.FontName = value;
                        }break;
                        default:
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                //throw new Xml2ExcelException(e.Message, e);
            }
            
        }


        private void GenerateColumn(Column column, IXLWorksheet xLWorksheet)
        {
            try
            {
                IXLColumn xLColumn = xLWorksheet.Column(column.number);
                if (column.width > 0)
                    xLColumn.Width = column.width;
                if (column.adjustToContents)
                    xLColumn.AdjustToContents();
                if (!string.IsNullOrEmpty(column.style))
                {
                    GenerateColumnStyles(xLColumn, column.style);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                //throw new Xml2ExcelException(e.Message, e);
            }
            
        }

        private void GenerateColumnStyles(IXLColumn xLColumn, string styles)
        {
            try
            {
                string[] styleList = styles.Split(';');
                var borderStyle = new Dictionary<string, int>()
                {
                    { "dotted",(int)XLBorderStyleValues.Dotted},
                    { "solid",(int)XLBorderStyleValues.Thick},
                    { "dashed",(int)XLBorderStyleValues.Dashed},
                    { "none",(int)XLBorderStyleValues.None},
                };
                var backgroundColorPattern = new Dictionary<string, int>()
                {
                    { "solid",(int)XLFillPatternValues.Solid},
                    { "darktrellis",(int)XLFillPatternValues.DarkTrellis},
                    { "lighttrellis",(int)XLFillPatternValues.LightTrellis},
                    { "darkhorizontal",(int)XLFillPatternValues.DarkHorizontal},
                    { "lighthorizontal",(int)XLFillPatternValues.LightHorizontal},
                    { "darkvertical",(int)XLFillPatternValues.DarkVertical},
                    { "lightvertical",(int)XLFillPatternValues.LightVertical},
                    { "darkdown",(int)XLFillPatternValues.DarkDown},
                    { "lightdown",(int)XLFillPatternValues.LightDown},
                    { "darkup",(int)XLFillPatternValues.DarkUp},
                    { "lightup",(int)XLFillPatternValues.LightUp},
                    { "lightgray",(int)XLFillPatternValues.LightGray},
                    { "darkgray",(int)XLFillPatternValues.DarkGray},
                    { "darkgrid",(int)XLFillPatternValues.DarkGrid},
                    { "lightgrid",(int)XLFillPatternValues.LightGrid},
                    { "none",(int)XLFillPatternValues.None},
                };
                foreach (var item in styleList)
                {
                    string[] keyValue = item.Split(':');
                    if (keyValue.Length != 2) continue;
                    string key = keyValue[0].Trim();
                    string value = keyValue[1].Trim();
                    switch (key)
                    {
                        case "text-align":
                            {
                                if (value == "left")
                                    xLColumn.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                                else if (value == "right")
                                    xLColumn.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                                else if (value == "center")
                                    xLColumn.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            }
                            break;
                        case "border-color":
                            {
                                xLColumn.Style.Border.SetLeftBorderColor(XLColor.FromHtml(value));
                                xLColumn.Style.Border.LeftBorder = (xLColumn.Style.Border.LeftBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLColumn.Style.Border.LeftBorder;

                                xLColumn.Style.Border.SetRightBorderColor(XLColor.FromHtml(value));
                                xLColumn.Style.Border.RightBorder = (xLColumn.Style.Border.RightBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLColumn.Style.Border.RightBorder;

                                xLColumn.Style.Border.SetTopBorderColor(XLColor.FromHtml(value));
                                xLColumn.Style.Border.TopBorder = (xLColumn.Style.Border.TopBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLColumn.Style.Border.TopBorder;

                                xLColumn.Style.Border.SetBottomBorderColor(XLColor.FromHtml(value));
                                xLColumn.Style.Border.BottomBorder = (xLColumn.Style.Border.BottomBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLColumn.Style.Border.BottomBorder;
                            }
                            break;
                        case "border-left-color":
                            {
                                xLColumn.Style.Border.SetLeftBorderColor(XLColor.FromHtml(value));
                                xLColumn.Style.Border.LeftBorder = (xLColumn.Style.Border.LeftBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLColumn.Style.Border.LeftBorder;

                            }
                            break;
                        case "border-right-color":
                            {
                                xLColumn.Style.Border.SetRightBorderColor(XLColor.FromHtml(value));
                                xLColumn.Style.Border.RightBorder = (xLColumn.Style.Border.RightBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLColumn.Style.Border.RightBorder;
                            }
                            break;
                        case "border-top-color":
                            {
                                xLColumn.Style.Border.SetTopBorderColor(XLColor.FromHtml(value));
                                xLColumn.Style.Border.TopBorder = (xLColumn.Style.Border.TopBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLColumn.Style.Border.TopBorder;
                            }
                            break;
                        case "border-bottom-color":
                            {
                                xLColumn.Style.Border.SetBottomBorderColor(XLColor.FromHtml(value));
                                xLColumn.Style.Border.BottomBorder = (xLColumn.Style.Border.BottomBorder == XLBorderStyleValues.None) ? XLBorderStyleValues.Thick : xLColumn.Style.Border.BottomBorder;
                            }
                            break;
                        case "border-style":
                            {
                                if (!borderStyle.ContainsKey(value)) continue;
                                borderStyle.TryGetValue(value, out int borderStyleValue);
                                xLColumn.Style.Border.TopBorder = (XLBorderStyleValues)borderStyleValue;
                                xLColumn.Style.Border.RightBorder = (XLBorderStyleValues)borderStyleValue;
                                xLColumn.Style.Border.BottomBorder = (XLBorderStyleValues)borderStyleValue;
                                xLColumn.Style.Border.LeftBorder = (XLBorderStyleValues)borderStyleValue;
                            }
                            break;
                        case "border-top-style":
                            {
                                if (!borderStyle.ContainsKey(value)) continue;
                                borderStyle.TryGetValue(value, out int borderStyleValue);
                                xLColumn.Style.Border.TopBorder = (XLBorderStyleValues)borderStyleValue;
                            }
                            break;
                        case "border-right-style":
                            {
                                if (!borderStyle.ContainsKey(value)) continue;
                                borderStyle.TryGetValue(value, out int borderStyleValue);
                                xLColumn.Style.Border.RightBorder = (XLBorderStyleValues)borderStyleValue;
                            }
                            break;
                        case "border-bottom-style":
                            {
                                if (!borderStyle.ContainsKey(value)) continue;
                                borderStyle.TryGetValue(value, out int borderStyleValue);
                                xLColumn.Style.Border.BottomBorder = (XLBorderStyleValues)borderStyleValue;
                            }
                            break;
                        case "border-left-style":
                            {
                                if (!borderStyle.ContainsKey(value)) continue;
                                borderStyle.TryGetValue(value, out int borderStyleValue);
                                xLColumn.Style.Border.LeftBorder = (XLBorderStyleValues)borderStyleValue;
                            }
                            break;
                        case "background-color":
                            {
                                xLColumn.Style.Fill.PatternType = (xLColumn.Style.Fill.PatternType == XLFillPatternValues.None) ? XLFillPatternValues.Solid : xLColumn.Style.Fill.PatternType;
                                xLColumn.Style.Fill.BackgroundColor = XLColor.FromHtml(value);
                            }
                            break;
                        case "background-color-pattern":
                            {
                                if (!backgroundColorPattern.ContainsKey(value)) continue;
                                backgroundColorPattern.TryGetValue(value, out int backgroundColorPatternValue);
                                xLColumn.Style.Fill.PatternType = (XLFillPatternValues)backgroundColorPatternValue;
                            }
                            break;
                        case "color":
                            {
                                xLColumn.Style.Font.FontColor = XLColor.FromHtml(value);
                            }
                            break;
                        case "font-style":
                            {
                                switch (value)
                                {
                                    case "italic": xLColumn.Style.Font.Italic = true; break;
                                    case "shadow": xLColumn.Style.Font.Shadow = true; break;
                                }
                            }
                            break;
                        case "font-weight":
                            {
                                switch (value)
                                {
                                    case "bold": xLColumn.Style.Font.SetBold(true); break;
                                    case "normal": xLColumn.Style.Font.Bold = false; break;
                                }
                            }
                            break;
                        case "text-decoration":
                            {
                                switch (value)
                                {
                                    case "none": xLColumn.Style.Font.SetUnderline(XLFontUnderlineValues.None); break;
                                    case "underline": xLColumn.Style.Font.SetUnderline(XLFontUnderlineValues.Single); break;
                                    case "underline-double": xLColumn.Style.Font.SetUnderline(XLFontUnderlineValues.Double); break;
                                    case "strikethrough": xLColumn.Style.Font.SetStrikethrough(true); break;
                                }
                            }
                            break;
                        case "font-size":
                            {
                                xLColumn.Style.Font.FontSize = double.Parse(value);
                            }
                            break;
                        case "font-family":
                            {
                                xLColumn.Style.Font.FontName = value;
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                //throw new Xml2ExcelException(e.Message, e);
            }
            
        }


    }
}
