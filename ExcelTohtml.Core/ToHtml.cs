using ExcelTohtml.Core.Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace ExcelTohtml.Core
{
    public class ToHtml
    {
        private readonly string TableStyle = "border-collapse: collapse;font-family: helvetica, arial, sans-serif;";
        private readonly bool DebugMode = false;

        private string ObjectJson;

        public readonly ExcelPackage Excel;
        private readonly ExcelWorksheet WorkSheet;
        private Dictionary<string, string> TemplateFieldList;
        private Dictionary<string, string> cellStyles;

        public ToHtml(FileInfo excelFile, string WorkSheetName = null)
        {
            if (!excelFile.Exists)
                throw new FileNotFoundException(string.Format("File {0} Not Found", excelFile.FullName));

            Excel = new ExcelPackage(excelFile);

            if (!string.IsNullOrEmpty(WorkSheetName))
                WorkSheet = Excel.Workbook.Worksheets[WorkSheetName];
            else
                WorkSheet = Excel.Workbook.Worksheets.First();
        }

        public ToHtml(ExcelWorksheet excelWorksheet)
        {
            WorkSheet = excelWorksheet ?? throw new ArgumentNullException(nameof(excelWorksheet));
        }

        /// <summary>
        /// Render HTML
        /// </summary>
        public string GetHtml()
        {
            //Check Performance
            var watch = System.Diagnostics.Stopwatch.StartNew();

            //GET DIMENSIONS
            var start = WorkSheet.Dimension.Start;
            var end = WorkSheet.Dimension.End;

            StringBuilder sb = new StringBuilder();

            // Row by row
            for (int row = start.Row; row <= end.Row; row++)
            {
                if (!WorkSheet.Row(row).Hidden)
                {
                    sb.AppendLine("<tr>");
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        if (!WorkSheet.Column(col).Hidden)
                        {
                            var d = WorkSheet.Cells[row, col];

                            int merged = 0;

                            if (d.Merge) //row is merged
                                merged = d.Worksheet.SelectedRange[WorkSheet.MergedCells[row, col]].Columns;

                            //11 default font size
                            var x = ProcessCellStyle(WorkSheet.Cells[row, col], WorkSheet.Column(col).Width, merged);
                            sb.AppendLine(x);
                            if (d.Merge)
                                col += (merged - 1);
                        }
                    }
                    sb.AppendLine("</tr>");
                }
            }

            sb.AppendLine("</table>");
            watch.Stop();
            var elapsedMs = watch.ElapsedMilliseconds;
            Console.WriteLine(" Processing time {0}ms", elapsedMs);

            return string.Format("<table  style=\"{0}>\" data-eth-ms=\"{1}\" data-eth-date=\"{2}\">{3}</table>",
                TableStyle, elapsedMs, DateTime.Now, sb.ToString());
        }

        private void IterateArray(JToken test)
        {
            foreach (var sub_obj in test)
            {
                if (sub_obj.Children().Any())
                {
                    IterateArray(sub_obj);
                }
                else
                {
                    Console.WriteLine("  {0}", sub_obj.Path);
                }
            }
        }

        private int CountElements(JToken test)
        {
            if (test == null)
                return 0;

            return test.Children().Count();
        }

        private void CopyRow(int rowFrom, int rows)
        {
            WorkSheet.InsertRow(rowFrom + 1, rows, rowFrom);
            var start = WorkSheet.Dimension.Start;
            var end = WorkSheet.Dimension.End;

            //iterate all
            for (int row = rowFrom + 1; row <= rowFrom + rows; row++)
            {
                for (int col = start.Column; col <= end.Column; col++)
                {
                    //copy from template
                    if (WorkSheet.Cells[rowFrom, col].Text.StartsWith("[["))
                        WorkSheet.Cells[row, col].Value = WorkSheet.Cells[rowFrom, col].Text.Replace("[!]", "[" + (row - rowFrom) + "]");

                    //We need to replicate EXcel Behavior for example
                    //Formula "=A1+1" will be stored in row1 = A1+1 in  row2  A2+1
                    if (!string.IsNullOrEmpty(WorkSheet.Cells[rowFrom, col].Formula))
                    {
                        WorkSheet.Cells[row, col].FormulaR1C1 = WorkSheet.Cells[rowFrom, col].FormulaR1C1;
                    }
                }
            }

            //fill initial row
            for (int col = start.Column; col <= end.Column; col++)
            {
                if (WorkSheet.Cells[rowFrom, col].Text.StartsWith("[["))
                    WorkSheet.Cells[rowFrom, col].Value = WorkSheet.Cells[rowFrom, col].Text.Replace("[!]", "[0]");
            }
        }

        public void DataFromObject(object data)
        {
            ObjectJson = JsonConvert.SerializeObject(data, Formatting.None);
        }

        public void DataFromUrl(string url)
        {
            Console.WriteLine("Connecting to {0}", url);

            using (WebClient wc = new WebClient())
            {
                ObjectJson = wc.DownloadString(url);
                ObjectJson = @"{""d"":[" + ObjectJson + "]}";
            }
            DataFromJson(ObjectJson);
        }

        public void DataFromJson(string objectJson)
        {
            JObject obj = JObject.Parse(objectJson);

            if (this.DebugMode)
                IterateArray(obj);

            var start = WorkSheet.Dimension.Start;
            var end = WorkSheet.Dimension.End;

            int _endRow = WorkSheet.Dimension.End.Row; // Template will extend umber of rows

            for (int row = start.Row; row <= _endRow; row++)
            {
                if (!WorkSheet.Row(row).Hidden)
                {
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        var d = WorkSheet.Cells[row, col];
                        string path = string.Empty;
                        int type = 0;
                        if (d.Text.StartsWith("[["))
                        {
                            type = 1;
                            path = d.Text.Replace("[[", "").Replace("]]", "");
                        }
                        else if (d.Comment != null && d.Comment.Text.StartsWith("[["))
                        {
                            type = 2;
                            path = d.Comment.Text.Replace("[[", "").Replace("]]", "");
                        }

                        if (type > 0) //found
                        {
                            //count items
                            if (path.Contains("[!]"))
                            {
                                string test = path.SubstringBefore("[!]");
                                JToken token1 = obj.SelectToken(test);
                                int rowsToCopy = CountElements(token1);
                                if (rowsToCopy > 0)// Spawn  rows
                                {
                                    _endRow += rowsToCopy;
                                    CopyRow(row, rowsToCopy - 1);

                                    if (type == 1)
                                    {
                                        path = d.Text.Replace("[[", "").Replace("]]", ""); //read one more time value changed
                                                                                           // row += rowsToCopy - 1; //skip created rows
                                    }
                                    else if (type == 2)
                                    {
                                        path = d.Comment.Text.Replace("[[", "").Replace("]]", ""); //read one more time value changed
                                                                                                   // row += rowsToCopy - 1; //skip created rows
                                    }
                                }
                            }

                            if (!path.Contains("[!]"))
                            {
                                JToken token = obj.SelectToken(path);

                                //if more than one for example array do nothing
                                if (token != null && !token.HasValues)
                                {
                                    if (decimal.TryParse(token.ToString(), out decimal myDec))
                                        d.Value = myDec;
                                    else
                                        d.Value = token.ToString();
                                }
                            }
                        }
                    }
                }
            }

            this.CalculateWorkbook();
        }

        /// <summary>
        /// Set Cell Values
        /// Supported Format
        /// A4 Test Value
        /// A5 =A2+A3
        /// [[TempalteField]] Test template
        /// </summary>
        public Dictionary<string, string> DataGetSet(Dictionary<string, string> data)
        {
            if (data == null)
                return new Dictionary<string, string>();

            //Dicionary to Excel
            foreach (var item in data)
            {
                //Template Handler
                if (item.Key.StartsWith("."))
                {
                    //output handler usefull for tests
                }
                else if (item.Key.StartsWith("[["))
                {
                    TemplateMap(); //One time

                    var FieldList = TemplateFieldList.Where(x => x.Value == item.Key);

                    foreach (var cellTemplate in FieldList)
                    {
                        if (item.Value.StartsWith("=")) //Formula
                            WorkSheet.Cells[cellTemplate.Key].Formula = item.Value.Remove(0, 1);
                        else //Text
                            WorkSheet.Cells[cellTemplate.Key].Value = item.Value;
                    }
                }
                else if (item.Value.StartsWith("=")) //Formula
                    WorkSheet.Cells[item.Key].Formula = item.Value.Remove(0, 1);
                else //Text
                    WorkSheet.Cells[item.Key].Value = item.Value;
            }

            this.CalculateWorkbook();

            string[] keys = data.Keys.ToArray();

            //Fill Out return values
            foreach (var item in keys)
            {
                if (item.StartsWith("."))
                {
                    data[item] = WorkSheet.Cells[item.Remove(0, 1)].Text;
                }
            }

            return data;
        }

        private void CalculateWorkbook()
        {
            if (Excel != null)
            {
                foreach (var _tempWorksheet in Excel.Workbook.Worksheets)
                    _tempWorksheet.Calculate();
            }
            else
                WorkSheet.Calculate();
        }

        /// <summary>
        /// Create Template Field Map, sample: [[Test]]
        /// </summary>
        private void TemplateMap()
        {
            if (TemplateFieldList == null)
            {
                TemplateFieldList = new Dictionary<string, string>();

                var start = WorkSheet.Dimension.Start;
                var end = WorkSheet.Dimension.End;

                for (int row = start.Row; row <= end.Row; row++)
                {
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        var cell = WorkSheet.Cells[row, col];
                        if (!String.IsNullOrEmpty(cell.Text) &&
                            cell.Text.StartsWith("[["))
                        {
                            TemplateFieldList.Add(cell.FullAddress, cell.Text);
                        }
                    }
                }
            }
        }

        private string ProcessCellStyle(ExcelRange input, double Width = -1, int ColSpan = 0)
        {
            cellStyles = new Dictionary<string, string>();

            StringBuilder sb = new StringBuilder();

            //Border
            PropertyToStyle("border-top", input.Style.Border.Top, cellAddress: input.Address);
            PropertyToStyle("border-right", input.Style.Border.Right, cellAddress: input.Address);
            PropertyToStyle("border-bottom", input.Style.Border.Bottom, cellAddress: input.Address);
            PropertyToStyle("border-left", input.Style.Border.Left, cellAddress: input.Address);

            //Align
            PropertyToStyle("text-align", input.Style.HorizontalAlignment.ToString(), "General");

            //Colors
            //Not properly implemented in Epplus  using ClosedXML

            PropertyToStyle("background-color", GetColor(input.Address, "background-color"));
            PropertyToStyle("color", GetColor(input.Address, "color"));

            PropertyToStyle("font-weight", input.Style.Font.Bold ? "bold" : "");
            PropertyToStyle("font-size", input.Style.Font.Size.ToString(), "11");
            PropertyToStyle("width", Convert.ToInt16(Width * 10));

            PropertyToStyle("white-space", !input.Style.WrapText ? "no-wrap" : "");

            string value = input.Text;
            if (string.IsNullOrEmpty(value))
                value = "&nbsp;";
            else
                value = System.Net.WebUtility.HtmlEncode(value);

            string comment = !string.IsNullOrEmpty(input.Comment?.Text) ? ("title=\"" + input.Comment.Text + "\"") : string.Empty;

            if (ColSpan > 0)
                sb.AppendFormat("<td style=\"{0}\" eth-cell=\"{1}\" colspan=\"{2}\" {4} >{3}</td>",
                    string.Join(";", cellStyles.Select(x => x.Key + ":" + x.Value)), input.Address, ColSpan, value, comment);
            else
                sb.AppendFormat("<td style=\"{0}\" eth-cell=\"{1}\" {3} >{2}</td>",
                    string.Join(";", cellStyles.Select(x => x.Key + ":" + x.Value)), input.Address, value, comment);

            return sb.ToString();
        }

        private void PropertyToStyle(string cssproperty, object value, string defaultValue = "", string cellAddress = "")
        {
            if (value == null)
                return;

            string cssItem;

            //borders
            if (value is ExcelBorderItem temp)
            {
                if (temp.Style == ExcelBorderStyle.None)
                    return;
                else if (temp.Style == ExcelBorderStyle.Thin)
                    cssItem = "solid 1px ";
                else if (temp.Style == ExcelBorderStyle.Hair || temp.Style == ExcelBorderStyle.Medium)
                    cssItem = "solid 2px ";
                else if (temp.Style == ExcelBorderStyle.Thick)
                    cssItem = "solid 3px ";
                else if (temp.Style == ExcelBorderStyle.Dashed)
                    cssItem = "dashed 1px ";
                else if (temp.Style == ExcelBorderStyle.Dotted)
                    cssItem = "dotted 1px ";
                else
                    cssItem = "solid 2px ";

                cssItem += GetColor(cellAddress, cssproperty);

                cellStyles.Add(cssproperty, cssItem);
                return;
            }
            else
            {
                cssItem = value.ToString();
            }

            if (cssItem != defaultValue)
            {
                if (cssproperty.Contains("size") || cssproperty.Contains("width"))
                {
                    cellStyles.Add(cssproperty, cssItem.Replace(",", ".") + "px");
                }
                else
                    cellStyles.Add(cssproperty, cssItem);
            }
        }

        private string GetColor(string address, string type)
        {
            var cell = WorkSheet.Cells[address];
            ExcelColor cellColor;
            if (type == "background-color")
                cellColor = cell.Style.Fill.BackgroundColor;
            else if (type == "color")
                cellColor = cell.Style.Font.Color;
            else if (type == "border-top")
                cellColor = cell.Style.Border.Top.Color;
            else if (type == "border-left")
                cellColor = cell.Style.Border.Left.Color;
            else if (type == "border-right")
                cellColor = cell.Style.Border.Right.Color;
            else if (type == "border-bottom")
                cellColor = cell.Style.Border.Bottom.Color;
            else
                return string.Empty;

            if (string.IsNullOrEmpty(cellColor.Rgb))
                return string.Empty;

            var intColor = int.Parse("FFFFFF00", System.Globalization.NumberStyles.HexNumber);
            var color = System.Drawing.Color.FromArgb(intColor);
            return System.Drawing.ColorTranslator.ToHtml(color);
        }
    }
}