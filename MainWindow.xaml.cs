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

namespace CRDEConverterJsonExcel;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
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
            string[] dialog = BrowseButton(sender, e, "json");
            string filePath = dialog[0]; // Ganti dengan path file JSON Anda
            string fileName = dialog[1]; // Ganti dengan path file JSON Anda
            string jsonContent = File.ReadAllText(filePath);
            // Parse JSON
            JObject jsonObject = JObject.Parse(jsonContent);
            JObject header = JObject.Parse(jsonContent);

            // Create Excel package
            using (var package = new ExcelPackage())
            {
                
                // Write data header
                ExcelWorksheet ws = package.Workbook.Worksheets.Add("#HEADER#");
                header.Value<JObject>("StrategyOneRequest").Value<JObject>("Body").Remove("Application_Header");
                ws.Cells[1, 1].Value = header.ToString();
                ws.Hidden = eWorkSheetHidden.VeryHidden;

                addSheet((JObject)jsonObject["StrategyOneRequest"]["Body"], package, null, 1, "-", 0);

                // Save Excel file
                DateTime timeStamp = DateTime.Now;
                string workingDirectory = Environment.CurrentDirectory;
                string projectDirectory = Directory.GetParent(workingDirectory).Parent.Parent.FullName;

                string fname = fileName + "-" + timeStamp.ToString("yyyyMMddHHmmssffff") + ".xlsx";
                string excelFilePath = projectDirectory + @"\output\excel\" + fname; // Ganti dengan path output Excel yang diinginkan
                package.SaveAs(new FileInfo(excelFilePath));

                MessageBox.Show("Conversion successful! File saved to " + excelFilePath);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error: " + ex.Message);
        }
    }

    private void btnConvertExcelToJSON_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            string[] dialog = BrowseButton(sender, e, "excel");
            string filePath = dialog[0];
            string fileName = dialog[1];

            // Set EPPlus license context (required for non-commercial use)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Read the Excel file
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var workbook = package.Workbook;
                JArray excelData = new JArray();
                //var excelData = new Dictionary<string, List<Dictionary<string, object>>>();

                for (int sheet = workbook.Worksheets.Count - 1; sheet >= 1; sheet--)
                {
                    var worksheet = workbook.Worksheets[sheet]; // Assuming the first sheet

                    if(worksheet.Dimension != null)
                    {
                        // Get the number of rows and columns
                        int rowCount = worksheet.Dimension.Rows;
                        int colCount = worksheet.Dimension.Columns;

                        // Create a list to hold the data

                        // Read the header row (first row)
                        var headers = new List<string>();
                        var typeDatas = new List<string>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            headers.Add(worksheet.Cells[1, col].Text);
                            typeDatas.Add(worksheet.Cells[2, col].Text);
                        }

                        //Data Kosong
                        if(rowCount < 3)
                        {
                            JObject emptyData = new JObject();
                            JObject cover = new JObject();
                            JObject variable = new JObject();
                            Int64 id;
                            Int64 parentId;

                            Int64.TryParse(worksheet.Cells[2, 1].Text, out id);
                            Int64.TryParse(worksheet.Cells[2, 1].Text, out parentId);

                            emptyData["Id"] = id;
                            emptyData["Parent"] = worksheet.Cells[2, 2].Text;
                            emptyData["ParentId"] = parentId;

                            variable["Variables"] = emptyData;
                            cover[worksheet.Name] = variable;
                            excelData.Add(cover);
                        }

                        // Read the data rows
                        for (int row = 3; row <= rowCount; row++) // Start from row 2 to skip header
                        {
                            var rowData = new JObject();
                            for (int col = 1; col <= colCount; col++)
                            {
                                string header = headers[col - 1];
                                string typeData = typeDatas[col - 1];
                                string cellValue = worksheet.Cells[row, col].Text;

                                double tempDouble;
                                Int64 tempInt;
                                if (cellValue == "")
                                {
                                    rowData[header] = cellValue;
                                } else
                                {
                                    switch (typeData)
                                    {
                                        case "Integer":
                                            Int64.TryParse(cellValue, out tempInt);
                                            rowData[header] = tempInt;
                                            break;
                                        case "Float":
                                            double.TryParse(cellValue, out tempDouble);
                                            rowData[header] = tempDouble;
                                            break;
                                        default:
                                            rowData[header] = cellValue;
                                            break;
                                    }
                                }
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

                //Mapping to JSON
                foreach (JObject data in excelData)
                {
                    foreach (var item in data)
                    {
                        JObject variable = (JObject)item.Value["Variables"];
                        Int64 idExcel;
                        Int64 parentIdExcel;
                        string parentExcel = variable["Parent"].ToString();

                        Int64.TryParse(variable["Id"].ToString(), out idExcel);
                        Int64.TryParse(variable["ParentId"].ToString(), out parentIdExcel);

                        // Clean Id, Parent, ParentId
                        variable.Remove("Id");
                        variable.Remove("Parent");
                        variable.Remove("ParentId");

                        if (parentExcel != null && parentExcel != "" && parentExcel != "-")
                        {
                            //JObject parent = excelData.Children<JObject>().FirstOrDefault(pnt => true);
                            JProperty parent = (JProperty)excelData.Children<JObject>().Children<JObject>().FirstOrDefault(pnt =>
                            {
                                JProperty parent = ((JProperty)pnt);
                                return parent.Name == parentExcel && parent.Value["Variables"]["Id"] != null && (int)parent.Value["Variables"]["Id"] == parentIdExcel;
                            });

                            JToken parentValue = parent.Value;
                            if (parentValue["Categories"] == null)
                            {
                                parentValue["Categories"] = new JArray();
                            }
                            ((JArray)parentValue["Categories"]).Add(data);
                        }
                    }
                }

                // Set Header JSON
                ExcelWorksheet ws = package.Workbook.Worksheets["#HEADER#"];
                string sheetHeader = ws.Cells[1, 1].Text;
                JObject headerJSON = JObject.Parse(sheetHeader);
                headerJSON["StrategyOneRequest"]["Body"] = excelData.Last;

                // Convert the data to JSON
                string json = JsonConvert.SerializeObject(headerJSON, Formatting.Indented);

                // Write the JSON to the specified file path
                DateTime timeStamp = DateTime.Now;
                string workingDirectory = Environment.CurrentDirectory;
                string projectDirectory = Directory.GetParent(workingDirectory).Parent.Parent.FullName;

                string fname = fileName + "-" + timeStamp.ToString("yyyyMMddHHmmssffff") + ".json";
                string jsonFilePath = projectDirectory + @"\output\json\" + fname; // Ganti dengan path output Excel yang diinginkan
                File.WriteAllText(jsonFilePath, json);

                MessageBox.Show("Excel file successfully converted to JSON and saved to: " + jsonFilePath);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error: {ex.Message}");
        }
    }

    private string[] BrowseButton(object sender, RoutedEventArgs e, string ext = "")
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

        OpenFileDialog dlg = new OpenFileDialog
        {
            Filter = filter,
        };

        // Set filter for file extension and default file extension 
        //dlg.DefaultExt = ".json";

        // Display OpenFileDialog by calling ShowDialog method 
        Nullable<bool> result = dlg.ShowDialog();

        // Get the selected file name and display in a TextBox 
        string filePath = "";
        string fileName = "";
        if (result == true)
        {
            // Open document 
            filePath = dlg.FileName;
            fileName = filePath.Split("\\").Last().Split(".").First();
        }

        return [filePath, fileName];
    }

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
                    {
                        row = startRow + 2;
                    }
                    else
                    {
                        valueStartRow = startRow - 1;
                    }

                    worksheet.Cells[2, 1].Value = "Integer";
                    worksheet.Cells[2, 2].Value = "String";
                    worksheet.Cells[2, 3].Value = "Integer";
                    worksheet.Cells[row, 1].Value = valueStartRow;
                    worksheet.Cells[row, 2].Value = parent;
                    worksheet.Cells[row, 3].Value = parentId;
                }

                // DictionaryHeader
                if (!dictionaryHeader.ContainsKey(worksheet.Name))
                {
                    dictionaryHeader.Add(worksheet.Name, new List<string>());
                }

                foreach (var variable in (JObject) property.Value)
                {
                    // Assign Dictionary Header
                    if (!dictionaryHeader[worksheet.Name].Contains(variable.Key))
                    {
                        dictionaryHeader[worksheet.Name].Add(variable.Key);
                    }

                    col = dictionaryHeader[worksheet.Name].IndexOf(variable.Key) + 4;
                    worksheet.Cells[1, col].Value = variable.Key;

                    // Coloring Header Backround Cell
                    worksheet.Cells[1, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[1, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.SkyBlue);

                    if (startRow == 1)
                    {
                        //Set Header
                        worksheet.Cells[row, 1].Value = "Integer";
                        worksheet.Cells[row, 2].Value = "String";
                        worksheet.Cells[row, 3].Value = "Integer";
                        worksheet.Cells[2, col].Value = variable.Value.Type;
                        row = startRow + 2;
                    }
                    else
                    {
                        valueStartRow = startRow - 1;
                    }
                    
                    worksheet.Cells[row, 1].Value = valueStartRow;
                    worksheet.Cells[row, 2].Value = parent;
                    worksheet.Cells[row, 3].Value = parentId;
                    worksheet.Cells[row, col].Value = (string) variable.Value;
                    col++;
                }
                worksheet.Row(2).Hidden = true;
                worksheet.Column(1).Hidden = true;
                worksheet.Column(2).Hidden = true;
                worksheet.Column(3).Hidden = true;
                worksheet.View.FreezePanes(2, 1);
            }
            else if (property.Key == "Categories")
            {
                foreach (var category in property.Value)
                {
                    addSheet((JObject)category, package, null, 1, parentName, startRow);
                }
            }
            else
            {
                if(package.Workbook.Worksheets[property.Key] == null)
                {
                    addSheet((JObject)property.Value, package, package.Workbook.Worksheets.Add(property.Key), 1, parent, parentId, property.Key);
                } else
                {
                    if(package.Workbook.Worksheets[property.Key].Dimension != null)
                    {
                        if (parentId > 1)
                        {
                            parentId -= 1;
                        }
                        addSheet((JObject)property.Value, package, package.Workbook.Worksheets[property.Key], package.Workbook.Worksheets[property.Key].Dimension.End.Row, parent, parentId, property.Key);
                    }
                }
            }
        }
    }
}