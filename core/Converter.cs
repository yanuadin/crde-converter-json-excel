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

namespace CRDEConverterJsonExcel.core
{
    class Converter
    {

        private Dictionary<string, List<string>> dictionaryHeader = new Dictionary<string, List<string>>();

        public JArray convertExcelTo(string fileName, string filePath, string convertType)
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
                for (int sheet = workbook.Worksheets.Count - 1; sheet >= 1; sheet--)
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
                        // Start from row 2 to skip header
                        for (int row = 3; row <= rowCount; row++)
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

                            //data.Add(rowData);
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
                        JObject variable = (JObject)item.Value["Variables"];
                        Int64 idExcel = convertTryParse(variable["Id"].ToString(), "Integer");
                        Int64 parentIdExcel = convertTryParse(variable["ParentId"].ToString(), "Integer");
                        string parentExcel = variable["Parent"].ToString();

                        if (parentExcel != null && parentExcel != "" && parentExcel != "-")
                        {
                            JProperty parent = (JProperty)excelData.Children<JObject>().Children<JObject>().FirstOrDefault(pnt =>
                            {
                                JProperty parent = (JProperty)pnt;
                                return parent.Name == parentExcel && parent.Value["Variables"]["Id"] != null && (int)parent.Value["Variables"]["Id"] == parentIdExcel;
                            });

                            JToken parentValue = parent.Value;
                            if (parentValue["Categories"] == null)
                                parentValue["Categories"] = new JArray();

                            ((JArray)parentValue["Categories"]).Add(data);
                        }
                        else
                        {
                            // Get Header JSOn
                            ExcelWorksheet ws = package.Workbook.Worksheets["#HEADER#"];
                            string sheetHeader = ws.Cells[(int)idExcel, 1].Text;
                            JObject headerJSON = JObject.Parse(sheetHeader);
                            result = new JObject();

                            try
                            {
                                // Set Header JSON
                                headerJSON["header"]["StrategyOneRequest"]["Body"] = excelData[iterator];

                                if (convertType == "json")
                                {
                                    // Convert the data to JSON
                                    jsonString = JsonConvert.SerializeObject(headerJSON["header"], Formatting.Indented);
                                    result["json"] = jsonString;
                                    result["fileName"] = headerJSON["name"];

                                    // Save the JSON file
                                    saveTextFile(@"\output\json\request\" + headerJSON["name"] + ".json", jsonString);
                                    result["message"] = @"[SUCCESS]: Request was saved in \output\json\request, please wait until response has been done!";
                                    result["success"] = true;

                                    resultCollection.Add(result);
                                }
                                else if (convertType == "txt")
                                {
                                    if (jsonString == "")
                                        jsonString = JsonConvert.SerializeObject(headerJSON["header"]);
                                    else
                                        jsonString += Environment.NewLine + JsonConvert.SerializeObject(headerJSON["header"]);

                                    result["message"] = @"[SUCCESS]: Excel file successfully converted and saved to \output\txt";
                                    result["success"] = true;
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

                // Clean Id, Parent, And ParentId

                try
                {
                    if (convertType == "txt")
                    {
                        // Save the JSON file
                        result["json"] = jsonString;
                        if (countApplicationHeader == 1)
                            result["fileName"] = fileName;
                        else
                            result["fileName"] = "MultipleFiles";

                        saveTextFile(@"\output\txt\" + result["fileName"] + ".txt", jsonString);

                        resultCollection.Add(result);
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

            return resultCollection;
        }

        private void cleanIdParentAndParentId()
        {

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
                        worksheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(CRDE.getColorCells[iterator]);
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

        public void saveTextFile(string filePath, string json)
        {
            // Arrange File Name
            string fileName = filePath.Split(@"\").Last().Split(".").First();
            string extension = filePath.Split(@"\").Last().Split(".").Last();
            string filePathWithoutName = string.Join(@"\", filePath.Split(@"\")[0..^1]) + @"\";

            string fname = fileName + "-" + GeneralMethod.getTimeStampNow() + "." + extension;
            string textFilePath = GeneralMethod.getProjectDirectory() + filePathWithoutName + fname;

            // Save Text File
            File.WriteAllText(textFilePath, json);
        }
    }
}
