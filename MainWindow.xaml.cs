using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using static OfficeOpenXml.ExcelErrorValue;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Reflection.PortableExecutable;
using System;
using System.Windows.Markup;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Reflection;
using CRDEConverterJsonExcel.core;
using System.Net.Http.Json;
using System.Xml.Linq;

namespace CRDEConverterJsonExcel;

public partial class MainWindow : Window
{
    Dictionary<string, List<string>> dictionaryHeader = new Dictionary<string, List<string>>();

    public MainWindow()
    {
        InitializeComponent();
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set license context for EPPlus
    }

    private void btnConvertJSONToExcel_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            // Create Excel package
            using (var package = new ExcelPackage())
            {
                JArray files = BrowseButton(sender, e, "json");

                // Arrange File Name
                string fname = (string) files.First["name"];
                if (files.Count > 1)
                    fname = "MultipleFiles";

                fname += "-" + getTimeStampNow() + ".xlsx";

                // Loop through the multiple files
                foreach (JObject file in files)
                {
                    string filePath = (string) file["path"];
                    string fileName = (string) file["name"];
                    string jsonContent = File.ReadAllText(filePath);

                    convertJSONToExcel(package, jsonContent, fileName);
                }

                // Save Excel file
                string excelFilePath = getProjectDirectory() + @"\output\excel\" + fname;
                package.SaveAs(new FileInfo(excelFilePath));

                MessageBox.Show(@"[SUCCESS] Conversion successful! File saved to \ouput\excel");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("[FAILED] Error: " + ex.Message);
        }
    }

    private void btnConvertExcelToJSON_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            // Browse for the Excel file
            JArray files = BrowseButton(sender, e, "excel");

            JObject result = convertExcelTo(files, "json");

            // Send Request to CRDE
            postRequestCRDE(result["json"].ToString(), result["fileName"].ToString());
        }
        catch (Exception ex)
        {
            MessageBox.Show($"[FAILED] Error: {ex.Message}");
        }
    }

    private void btnConvertExcelToTxt_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            // Browse for the Excel file
            JArray files = BrowseButton(sender, e, "excel");

            convertExcelTo(files, "txt");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"[FAILED] Error: {ex.Message}");
        }
    }

    private void convertJSONToExcel(ExcelPackage package, string json, string fileName)
    {
        // Parse JSON
        JObject jsonObject = JObject.Parse(json);
        JObject header = JObject.Parse(json);

        // Write data header
        ExcelWorksheet ws = package.Workbook.Worksheets["#HEADER#"];
        int rowHeader = 1;
        if (ws == null)
            ws = package.Workbook.Worksheets.Add("#HEADER#");
        else
            rowHeader = ws.Dimension.End.Row + 1;

        // Remove Application Header
        JObject hdr = (JObject)header.First.First.Last.First;
        hdr.Remove("Application_Header");

        JObject headerJSON = new JObject();
        headerJSON["name"] = fileName;
        headerJSON["header"] = header;
        ws.Cells[rowHeader, 1].Value = headerJSON.ToString();
        ws.Hidden = eWorkSheetHidden.VeryHidden;

        // Start Recursive Looping with parameter Application Header as JObject
        addSheet((JObject)jsonObject.First.First.Last.First, package, null, 1, "-", 0);
    }

    private JObject convertExcelTo(JArray files, string convertType)
    {
        string message = "";
        JObject result = new JObject();

        foreach (JObject file in files)
        {
            string filePath = (string)file["path"];
            string fileName = (string)file["name"];

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
                foreach (JObject data in excelData)
                {
                    foreach (var item in data)
                    {
                        JObject variable = (JObject)item.Value["Variables"];
                        Int64 idExcel = convertTryParse(variable["Id"].ToString(), "Integer");
                        Int64 parentIdExcel = convertTryParse(variable["ParentId"].ToString(), "Integer");
                        string parentExcel = variable["Parent"].ToString();

                        // Clean Id, Parent, ParentId
                        variable.Remove("Id");
                        variable.Remove("Parent");
                        variable.Remove("ParentId");

                        if (parentExcel != null && parentExcel != "" && parentExcel != "-")
                        {
                            //JObject parent = excelData.Children<JObject>().FirstOrDefault(pnt => true);
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
                            // Set Header JSON
                            ExcelWorksheet ws = package.Workbook.Worksheets["#HEADER#"];
                            string sheetHeader = ws.Cells[(int)idExcel, 1].Text;
                            JObject headerJSON = JObject.Parse(sheetHeader);
                            headerJSON["header"]["StrategyOneRequest"]["Body"] = excelData[iterator];

                            if (convertType == "json")
                            {
                                // Convert the data to JSON
                                jsonString = JsonConvert.SerializeObject(headerJSON["header"], Formatting.Indented);
                                result["json"] = jsonString;
                                result["fileName"] = headerJSON["name"];

                                // Save the JSON file
                                saveTextFile(@"\output\json\request\" + headerJSON["name"] + ".json", jsonString);
                                message = @"[SUCCESS] Request was saved in \output\json\request, please wait until response has been done!";
                            }
                            else if (convertType == "txt")
                            {
                                if (jsonString == "")
                                    jsonString = JsonConvert.SerializeObject(headerJSON["header"]);
                                else
                                    jsonString += Environment.NewLine + JsonConvert.SerializeObject(headerJSON["header"]);
                            }
                            else
                            {
                                message = "[FAILED] Invalid Convert Type";
                                break;
                            }

                            countApplicationHeader++;
                        }
                    }
                    iterator++;
                }

                if (convertType == "txt")
                {
                    // Save the JSON file
                    result["json"] = jsonString;
                    if(countApplicationHeader == 1)
                        result["fileName"] = fileName;
                    else
                        result["fileName"] = "MultipleFiles";

                    saveTextFile(@"\output\txt\"+ result["fileName"] + ".txt", jsonString);

                    message = @"[SUCCESS] Excel file successfully converted and saved to \output\txt";
                }
            }
        }

        MessageBox.Show(message);

        return result;
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

    private JArray BrowseButton(object sender, RoutedEventArgs e, string ext = "")
    {
        // Create OpenFileDialog 
        string filter = "";
        switch (ext)
        {
            case "json":
                filter = "Json files (*.json)|*.json";
                break;
            case "excel":
                filter = "Excel Files|*.xls;*.xlsx";
                break;
            default:
                filter = "Json files (*.json)|*.json|Excel Files|*.xls;*.xlsx";
                break;
        }

        OpenFileDialog dlg = new OpenFileDialog { Filter = filter, Multiselect = true };

        // Display OpenFileDialog by calling ShowDialog method 
        Nullable<bool> result = dlg.ShowDialog();

        // Get the selected file name and display in a TextBox 
        string filePath = "";
        string fileName = "";
        JArray files = new JArray();
        if (result == true)
        {
            // Open document 
            foreach (string file in dlg.FileNames)
            {
                JObject fileProperties = new JObject();
                fileProperties["path"] = file;
                fileProperties["name"] = file.Split("\\").Last().Split(".").First();
                files.Add(fileProperties);
            }
        }

        return files;
    }

    // Recursive Looping
    private void addSheet(JObject data, ExcelPackage package, ExcelWorksheet worksheet = null, int startRow = 1, string parent = "", int parentId = 1, string parentName = "")
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
                        row = startRow + 2;
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

                foreach (var variable in (JObject) property.Value)
                {
                    // Assign Dictionary Header
                    if (!dictionaryHeader[worksheet.Name].Contains(variable.Key))
                        dictionaryHeader[worksheet.Name].Add(variable.Key);

                    col = dictionaryHeader[worksheet.Name].IndexOf(variable.Key) + 4;
                    worksheet.Cells[1, col].Value = variable.Key;

                    // Coloring Header Backround Cell
                    worksheet.Cells[1, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[1, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.SkyBlue);

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
                    addSheet((JObject)category, package, null, 1, parentName, startRow);
            }
            else
            {
                if(package.Workbook.Worksheets[property.Key] == null)
                    addSheet((JObject)property.Value, package, package.Workbook.Worksheets.Add(property.Key), 1, parent, parentId, property.Key);
                else
                {
                    if(package.Workbook.Worksheets[property.Key].Dimension != null)
                    {
                        if (parentId > 1)
                            parentId -= 1;

                        addSheet((JObject)property.Value, package, package.Workbook.Worksheets[property.Key], package.Workbook.Worksheets[property.Key].Dimension.End.Row, parent, parentId, property.Key);
                    }
                }
            }
        }
    }

    private void saveTextFile(string filePath, string json)
    {
        // Arrange File Name
        string fileName = filePath.Split(@"\").Last().Split(".").First();
        string extension = filePath.Split(@"\").Last().Split(".").Last();
        string filePathWithoutName = string.Join(@"\", filePath.Split(@"\")[0..^1]) + @"\";

        string fname = fileName + "-" + getTimeStampNow() + "." + extension;
        string textFilePath = getProjectDirectory() + filePathWithoutName + fname;

        // Save Text File
        File.WriteAllText(textFilePath, json);
    }

    private string getProjectDirectory()
    {
        string workingDirectory = Environment.CurrentDirectory;
        string projectDirectory = Directory.GetParent(workingDirectory).Parent.Parent.FullName;

        return projectDirectory;
    }

    private string getTimeStampNow()
    {
        return DateTime.Now.ToString("yyyyMMddHHmmssffff");
    }

    private async void postRequestCRDE(string json, string saveFileNameResponse)
    {
        saveFileNameResponse = saveFileNameResponse + "_response";
        // API endpoint
        string apiUrl = "https://crde-etl-uat.mylab.local/api/v1/s1/online";

        // Parse JSON
        JObject jsonObject = JObject.Parse(json);

        // Data to send in the POST request
        try
        {
            using (var package = new ExcelPackage())
            {
                // Call the API and get the response
                string responseJsonText = await Api.PostApiDataAsync(apiUrl, jsonObject);
                JObject parseResponseJson = JObject.Parse(responseJsonText);
                string responseJsonIndent = JsonConvert.SerializeObject(parseResponseJson, Formatting.Indented);

                // Save Response to JSON File
                saveTextFile(@"\output\json\response\" + saveFileNameResponse + ".json", responseJsonIndent);
                MessageBox.Show(@"[SUCCESS] Response was saved in \output\json\response");

                // Convert Response to Excel
                convertJSONToExcel(package, responseJsonText, saveFileNameResponse);

                // Save Excel file
                string excelFilePath = getProjectDirectory() + @"\output\excel\" + saveFileNameResponse + '-' + getTimeStampNow() + ".xlsx";
                package.SaveAs(new FileInfo(excelFilePath));

                MessageBox.Show("[SUCCESS] " + saveFileNameResponse + @" Save Response was successful! File saved to \output\excel");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"[API_FAILED] An error occurred: {ex.Message}", "Error");
        }
    }
}