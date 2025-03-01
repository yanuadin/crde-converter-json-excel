using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Diagnostics;
using CRDEConverterJsonExcel.config;
using System.Drawing;

namespace CRDEConverterJsonExcel.core
{
    class Converter
    {
        private Dictionary<string, List<string>> dictionaryHeader = new Dictionary<string, List<string>>();
        private CRDE config = new CRDE();

        public void convertJSONToExcel(ExcelPackage package, string json, int iterator)
        {
            // Parse JSON
            JObject jsonObject = JObject.Parse(json);
            JObject header = JObject.Parse(json);

            // Write data header
            ExcelWorksheet ws = package.Workbook.Worksheets["#HEADER#"];
            int rowHeader = 3;
            if (ws == null)
                ws = package.Workbook.Worksheets.Add("#HEADER#");
            else
                rowHeader = ws.Dimension.End.Row + 1;

            // Remove Application Header
            JObject hdr = (JObject)header.First.First.Last.First;
            hdr.Remove("Application_Header");

            // Set Header To Sheet
            ws.Cells[1, 1].Value = "Id";
            ws.Cells[1, 2].Value = "Parent";
            ws.Cells[1, 3].Value = "ParentId";
            
            ws.Cells[2, 1].Value = "Integer";
            ws.Cells[2, 2].Value = "String";
            ws.Cells[2, 3].Value = "Integer";

            // Add Dictionary Header Header
            if (!dictionaryHeader.ContainsKey("#HEADER#"))
                dictionaryHeader.Add("#HEADER#", new List<string>());

            // Type Data Header
            dictionaryHeader["#HEADER#"].Add("Type");
            int colType = dictionaryHeader["#HEADER#"].IndexOf("Type") + 4;
            ws.Cells[1, colType].Value = "Type";
            ws.Cells[2, colType].Value = "String";
            ws.Cells[rowHeader, colType].Value = header.First.ToObject<JProperty>().Name;

            // Coloring Type Data Header Background Cell
            ws.Cells[1, colType].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Cells[1, colType].Style.Fill.BackgroundColor.SetColor(Color.Silver);

            // Coloring Type Data Background Cell
            ws.Cells[rowHeader, colType].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Cells[rowHeader, colType].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(config.getColorCells()[iterator].ToString()));

            foreach (var prop in (JObject) header.First.First["Header"])
            {
                // Assign Dictionary Header Header
                if (!dictionaryHeader["#HEADER#"].Contains(prop.Key))
                {
                    int colHeader = dictionaryHeader["#HEADER#"].Count() + 4;
                    ws.Cells[1, colHeader].Value = prop.Key;
                    ws.Cells[2, colHeader].Value = prop.Value.Type;
                    dictionaryHeader["#HEADER#"].Add(prop.Key);

                    // Coloring Header Background Cell
                    ws.Cells[1, colHeader].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    ws.Cells[1, colHeader].Style.Fill.BackgroundColor.SetColor(Color.Silver);
                }

                int col = dictionaryHeader["#HEADER#"].IndexOf(prop.Key) + 4;
                
                ws.Cells[rowHeader, 1].Value = iterator + 1;
                ws.Cells[rowHeader, 2].Value = "-";
                ws.Cells[rowHeader, 3].Value = 0;
                ws.Cells[rowHeader, col].Value = prop.Value.ToString();

                // Coloring  Background Cell
                ws.Cells[rowHeader, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws.Cells[rowHeader, col].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(config.getColorCells()[iterator].ToString()));
            }

            // Hide Parent Child Pointer and Freeze Header
            ws.Row(2).Hidden = true;
            ws.Column(1).Hidden = true;
            ws.Column(2).Hidden = true;
            ws.Column(3).Hidden = true;
            ws.View.FreezePanes(2, 1);

            // Start Recursive Looping with parameter Application Header as JObject
            addSheet(iterator, (JObject)jsonObject.First.First.Last.First, package, null, 1, "#HEADER#", iterator + 2);
        }

        public JArray convertExcelTo(string filePath, string convertType)
        {
            JArray resultCollection = new JArray();

            // Set EPPlus license context (required for non-commercial use)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Read the Excel file
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var workbook = package.Workbook;
                JArray excelData = new JArray();

                // Loop through the worksheets in the Excel file to JSON
                for (int sheet = workbook.Worksheets.Count - 1; sheet >= 0; sheet--)
                {
                    // Get the worksheet by name
                    var worksheet = workbook.Worksheets[sheet];

                    if (worksheet.Dimension != null)
                    {
                        // Get the number of rows and columns
                        int rowCount = worksheet.Dimension.Rows;
                        int colCount = worksheet.Dimension.Columns;

                        // Read the header row (first row)
                        var headers = new List<string>();
                        var typeDatas = new List<string>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            headers.Add(worksheet.Cells[1, col].Text);
                            typeDatas.Add(worksheet.Cells[2, col].Text);
                        }

                        // Empty Data
                        if (rowCount < 3)
                        {
                            JObject emptyData = new JObject();
                            JObject cover = new JObject();
                            JObject variable = new JObject();

                            emptyData["Id"] = convertTryParse(worksheet.Cells[2, 1].Text, "Integer");
                            emptyData["Parent"] = worksheet.Cells[2, 2].Text;
                            emptyData["ParentId"] = convertTryParse(worksheet.Cells[2, 1].Text, "Integer");
                            variable["Variables"] = emptyData;
                            cover[worksheet.Name] = variable;
                            excelData.Add(cover);
                        }

                        // Read the data rows
                        // Start from row 3 to skip header
                        for (int row = rowCount; row >= 3; row--)
                        {
                            var rowData = new JObject();
                            for (int col = 1; col <= colCount; col++)
                            {
                                string header = headers[col - 1];
                                string typeData = typeDatas[col - 1];
                                string cellValue = worksheet.Cells[row, col].Text;

                                if (cellValue == "")
                                    rowData[header] = cellValue;
                                else
                                    rowData[header] = convertTryParse(cellValue, typeData);
                            }

                            JObject cover = new JObject();
                            JObject variable = new JObject();
                            variable["Variables"] = rowData;
                            cover[worksheet.Name] = variable;
                            excelData.Add(cover);
                        }
                    }
                }

                //Mapping Children to Parent
                int iterator = 0;
                string jsonString = "";
                int countApplicationHeader = 0;
                JObject result = new JObject();
                foreach (JObject data in excelData)
                {
                    foreach (var item in data)
                    {
                        JObject variable = item.Key == "#HEADER#" ? (JObject)item.Value.First.First.First.First : (JObject)item.Value["Variables"];
                        Int64 idExcel = convertTryParse(variable["Id"].ToString(), "Integer");
                        Int64 parentIdExcel = convertTryParse(variable["ParentId"].ToString(), "Integer");
                        string parentExcel = variable["Parent"].ToString();

                        if (parentExcel != null && parentExcel != "" && parentExcel != "-")
                        {
                            var parent = excelData.Children<JObject>().Children<JObject>().FirstOrDefault(pnt =>
                            {
                                JProperty parent = (JProperty)pnt;
                                return parent.Value["Variables"] != null && parent.Name == parentExcel && parent.Value["Variables"]["Id"] != null && (int)parent.Value["Variables"]["Id"] == parentIdExcel;
                            });

                            JProperty parentProperty = (JProperty)parent;
                            if (parentProperty.Name == "#HEADER#")
                            {
                                JObject skletonTypeHeader = new JObject();
                                JObject skletonHeader = new JObject();

                                skletonHeader["Header"] = parentProperty.Value["Variables"];
                                skletonHeader["Body"] = data;
                                skletonTypeHeader[parentProperty.Value["Variables"]["Type"].ToString()] = skletonHeader;
                                parentProperty.Value = skletonTypeHeader;
                            }
                            else
                            {
                                if (parentProperty.Value["Categories"] == null)
                                    parentProperty.Value["Categories"] = new JArray();

                                ((JArray)parentProperty.Value["Categories"]).AddFirst(data);
                            }
                        }
                        else
                        {
                            JObject headerJSON = new JObject();
                            headerJSON["header"] = cleanIdParentAndParentId(excelData[iterator]["#HEADER#"].ToObject<JObject>());
                            headerJSON["name"] = headerJSON["header"].First.First.First.First["InquiryCode"];
                            result = new JObject();

                            try
                            {
                                if (convertType == "json")
                                {
                                    // Convert the data to JSON
                                    result["json"] = JsonConvert.SerializeObject(headerJSON["header"], Formatting.Indented);
                                    result["fileName"] = headerJSON["name"];
                                    result["message"] = @"[SUCCESS]: Request was saved in \output\json\request";
                                    result["success"] = true;

                                    resultCollection.Add(result);
                                }
                                else if (convertType == "txt")
                                {
                                    result["json"] = JsonConvert.SerializeObject(headerJSON["header"]);
                                    result["fileName"] = headerJSON["name"];
                                    result["message"] = @"[SUCCESS]: Request was saved in \output\txt";
                                    result["success"] = true;

                                    resultCollection.Add(result);
                                }
                                else
                                {
                                    result["json"] = "";
                                    result["fileName"] = "";
                                    result["message"] = "[FAILED]: Invalid Convert Type";
                                    result["success"] = false;

                                    break;
                                }
                            }
                            catch (Exception ex)
                            {
                                result["json"] = "";
                                result["fileName"] = "";
                                result["message"] = "[FAILED]: [" + headerJSON["name"] + "] Convert was failed";
                                result["success"] = false;

                                continue;
                            }

                            countApplicationHeader++;
                        }
                        iterator++;
                    }
                }

                // Save File
                foreach (JObject res in resultCollection)
                {
                    try
                    {
                        if (convertType == "json")
                        {
                            saveTextFile(@"\output\json\request\" + res["fileName"] + ".json", res["json"].ToString());
                        } 
                        else if (convertType == "txt")
                        {
                            if (res.Previous != null)
                            {
                                resultCollection[0]["json"] += Environment.NewLine + res["json"].ToString();
                                resultCollection[0]["fileName"] = "MultipleFiles";
                            }

                            if (res.Next == null)
                                saveTextFile(@"\output\txt\" + resultCollection[0]["fileName"] + ".txt", resultCollection[0]["json"].ToString());
                        }
                    }
                    catch (Exception ex)
                    {
                        result["json"] = "";
                        result["fileName"] = "";
                        result["message"] = "[FAILED]: Convert was failed";
                        result["success"] = false;
                    }
                }
            }

            return resultCollection;
        }
        private JObject cleanIdParentAndParentId(JObject data)
        {
            foreach (var property in data)
            {
                if (property.Value.GetType().ToString() == "Newtonsoft.Json.Linq.JObject" )
                {
                    if (property.Key == "Variables")
                    {
                        JObject variable = (JObject)property.Value;
                        variable.Remove("Id");
                        variable.Remove("Parent");
                        variable.Remove("ParentId");
                        data["Variables"] = variable;
                    } else if (property.Key == "Header")
                    {
                        JObject variable = (JObject)property.Value;
                        variable.Remove("Id");
                        variable.Remove("Parent");
                        variable.Remove("ParentId");
                        variable.Remove("Type");
                        data[property.Key] = variable;
                    }
                    else
                        cleanIdParentAndParentId((JObject)property.Value);
                }
                else if (property.Value.GetType().ToString() == "Newtonsoft.Json.Linq.JArray" && property.Key == "Categories")
                    foreach (var category in property.Value)
                        cleanIdParentAndParentId((JObject)category);
            }

            return data;
        }

        private dynamic convertTryParse(dynamic value, string typeData)
        {
            double tempDouble;
            Int64 tempInt;
            DateTime tempDateTime;
            dynamic result;

            switch (typeData)
            {
                case "Integer":
                    Int64.TryParse(value, out tempInt);
                    result = tempInt;
                    break;
                case "Float":
                    double.TryParse(value, out tempDouble);
                    result = tempDouble;
                    break;
                case "Date":
                    DateTime.TryParse(value, out tempDateTime);
                    result = tempDateTime.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss");
                    break;
                default:
                    result = value;
                    break;
            }

            return result;
        }

        // Recursive Looping
        public void addSheet(int iterator, JObject data, ExcelPackage package, ExcelWorksheet worksheet = null, int startRow = 1, string parent = "", int parentId = 1, string parentName = "")
        {
            foreach (var property in data)
            {
                //Assign to Excel
                if (property.Key == "Variables")
                {
                    int col = 4;
                    int row = startRow + 1;
                    int valueStartRow = startRow;

                    worksheet.Cells[1, 1].Value = "Id";
                    worksheet.Cells[1, 2].Value = "Parent";
                    worksheet.Cells[1, 3].Value = "ParentId";
                    if (property.Value.Count() == 0)
                    {
                        if (startRow == 1)
                        {
                            row = startRow + 2;
                        }
                        else
                            valueStartRow = startRow - 1;

                        worksheet.Cells[2, 1].Value = "Integer";
                        worksheet.Cells[2, 2].Value = "String";
                        worksheet.Cells[2, 3].Value = "Integer";
                        worksheet.Cells[row, 1].Value = valueStartRow;
                        worksheet.Cells[row, 2].Value = parent;
                        worksheet.Cells[row, 3].Value = parentId;
                    }

                    // DictionaryHeader
                    if (!dictionaryHeader.ContainsKey(worksheet.Name))
                        dictionaryHeader.Add(worksheet.Name, new List<string>());

                    foreach (var variable in (JObject)property.Value)
                    {
                        // Assign Dictionary Header
                        if (!dictionaryHeader[worksheet.Name].Contains(variable.Key))
                            dictionaryHeader[worksheet.Name].Add(variable.Key);

                        col = dictionaryHeader[worksheet.Name].IndexOf(variable.Key) + 4;
                        worksheet.Cells[1, col].Value = variable.Key;

                        // Coloring Header Background Cell
                        worksheet.Cells[1, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[1, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Silver);

                        if (startRow == 1)
                        {
                            // Set Header
                            worksheet.Cells[row, 1].Value = "Integer";
                            worksheet.Cells[row, 2].Value = "String";
                            worksheet.Cells[row, 3].Value = "Integer";
                            worksheet.Cells[2, col].Value = variable.Value.Type;
                            row = startRow + 2;
                        }
                        else
                            valueStartRow = startRow - 1;

                        // Re-Check Type If Empty Cell
                        if (variable.Value.Type.ToString() != worksheet.Cells[2, col].Text && worksheet.Cells[2, col].Text == "String")
                            worksheet.Cells[2, col].Value = variable.Value.Type;

                        worksheet.Cells[row, 1].Value = valueStartRow;
                        worksheet.Cells[row, 2].Value = parent;
                        worksheet.Cells[row, 3].Value = parentId;
                        worksheet.Cells[row, col].Value = variable.Value.ToString();

                        // Coloring  Background Cell
                        worksheet.Cells[row, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(config.getColorCells()[iterator].ToString()));
                    }

                    // Hide Parent Child Pointer and Freeze Header
                    worksheet.Row(2).Hidden = true;
                    worksheet.Column(1).Hidden = true;
                    worksheet.Column(2).Hidden = true;
                    worksheet.Column(3).Hidden = true;
                    worksheet.View.FreezePanes(2, 1);
                }
                else if (property.Key == "Categories")
                {
                    foreach (var category in property.Value)
                        addSheet(iterator, (JObject)category, package, null, 1, parentName, startRow);
                }
                else
                {

                    if (parentId > 1)
                        parentId -= 1;

                    if (package.Workbook.Worksheets[property.Key] == null)
                        addSheet(iterator, (JObject)property.Value, package, package.Workbook.Worksheets.Add(property.Key), 1, parent, parentId, property.Key);
                    else
                    {
                        if (package.Workbook.Worksheets[property.Key].Dimension != null)
                        {

                            addSheet(iterator, (JObject)property.Value, package, package.Workbook.Worksheets[property.Key], package.Workbook.Worksheets[property.Key].Dimension.End.Row, parent, parentId, property.Key);
                        }
                    }
                }
            }
        }

        public void saveTextFile(string filePath, string json, string typeJSON = "req")
        {
            // Arrange File Name
            string fileName = filePath.Split(@"\").Last().Split(".").First();
            string extension = filePath.Split(@"\").Last().Split(".").Last();
            string filePathWithoutName = string.Join(@"\", filePath.Split(@"\")[0..^1]) + @"\";

            string fname = fileName + "-" + typeJSON + "-" + GeneralMethod.getTimeStampNow() + "." + extension;
            string textFilePath = GeneralMethod.getProjectDirectory() + filePathWithoutName + fname;

            // Save Text File
            File.WriteAllText(textFilePath, json);
        }
    }
}
