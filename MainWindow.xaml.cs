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
using System.Collections;
using CRDEConverterJsonExcel.config;
using System.Data;

namespace CRDEConverterJsonExcel;

public partial class MainWindow : Window
{
    Converter converter = new Converter();
    ConverterV2 converterV2 = new ConverterV2();
    List<Item> lb_requestItems = new List<Item>();

    // Define a class to represent each item in the ListBox
    public class Item
    {
        public string fileName { get; set; }
        public string json { get; set; }
        public bool isSelected { get; set; }
    }

    public MainWindow()
    {
        InitializeComponent();
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set license context for EPPlus

        // Initialize Endpoint Combobox
        foreach (string endpoint in CRDE.getAllEndpoint())
            cb_endpoint.Items.Add(endpoint);
    }

    private void btnConvertJSONToExcel_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            // Create Excel package
            using (var package = new ExcelPackage())
            {
                JArray files = BrowseButton(sender, e, "json", true);

                // Arrange File Name
                string fname = files.First["name"].ToString();
                if (files.Count > 1)
                    fname = "MultipleFiles";

                fname += "-" + GeneralMethod.getTimeStampNow() + ".xlsx";

                // Loop through the multiple files
                int iterator = 0;
                foreach (JObject file in files)
                {
                    string filePath = file["path"].ToString();
                    string fileName = file["name"].ToString();
                    string jsonContent = File.ReadAllText(filePath);

                    convertJSONToExcel(package, jsonContent, fileName, iterator++);
                }

                // Save Excel file
                string excelFilePath = GeneralMethod.getProjectDirectory() + @"\output\excel\request\" + fname;
                package.SaveAs(new FileInfo(excelFilePath));

                MessageBox.Show(@"[SUCCESS]: Conversion successful! File saved to \ouput\excel\request");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("[FAILED]: Error: " + ex.Message);
        }
    }

    private void btnConvertJSONToExcelV2_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            // Create Excel package
            using (var package = new ExcelPackage())
            {
                JObject file = BrowseButton(sender, e, "json", false);
                string filePath = file["path"].ToString();
                string fileName = file["name"].ToString();
                string jsonContent = File.ReadAllText(filePath);

                convertJSONToExcelV2(package, jsonContent, fileName);

                // Save Excel file
                string excelFilePath = GeneralMethod.getProjectDirectory() + @"\output\excel\" + fileName + "-" + GeneralMethod.getTimeStampNow() + ".xlsx";
                package.SaveAs(new FileInfo(excelFilePath));

                MessageBox.Show(@"[SUCCESS]: Conversion successful! File saved to \ouput\excel");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("[FAILED]: Error: " + ex.Message);
        }
    }

    private void btnConvertExcelToJSON_Click(object sender, RoutedEventArgs e)
    {
        // Browse for the Excel file
        JObject file = BrowseButton(sender, e, "excel", false);
        string filePath = file["path"].ToString();
        string fileName = file["name"].ToString();

        JArray result = new JArray();

        // Convertin to JSON
        result = converter.convertExcelTo(fileName, filePath, "json");

        // Bind the list to the ListBox
        int successCount = 0;
        int errorCount = 0;
        lb_requestItems = new List<Item>();
        foreach (JObject res in result)
        {
            if (bool.Parse(res["success"].ToString()))
            {
                successCount++;
                lb_requestItems.Add(new Item { fileName = res["fileName"].ToString(), json = res["json"].ToString(), isSelected = false });
            }
            else 
                errorCount++;
        }
        lb_requestList.ItemsSource = lb_requestItems;

        // Print Message Box
        MessageBox.Show($"[SUCCESS]: {successCount} files converted successfully, {errorCount} files failed to convert" + Environment.NewLine + Environment.NewLine + "File was saved in " + @"\output\json\request");
    }

    private void btnConvertExcelToTxt_Click(object sender, RoutedEventArgs e)
    {
        // Browse for the Excel file
        JObject file = BrowseButton(sender, e, "excel", false);
        string filePath = file["path"].ToString();
        string fileName = file["name"].ToString();

        JArray result = new JArray();

        // Convertin to JSON
        result = converter.convertExcelTo(fileName, filePath, "txt");

        // Print Message Box
        int successCount = 0;
        int errorCount = 0;
        foreach (JObject res in result)
        {
            if (bool.Parse(res["success"].ToString()))
                successCount++;
            else
                errorCount++;
        }

        //converter.convertExcelTo(files, "txt");

        MessageBox.Show($"[SUCCESS]: {successCount} files converted successfully, {errorCount} files failed to convert" + Environment.NewLine + Environment.NewLine + "File was saved in " + @"\output\json\request");
    }

    private void convertJSONToExcel(ExcelPackage package, string json, string fileName, int iterator)
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
        converter.addSheet(iterator, (JObject)jsonObject.First.First.Last.First, package, null, 1, "-", 0);
    }

    private void convertJSONToExcelV2(ExcelPackage package, string json, string fileName)
    {
        // Parse JSON
        JObject jsonObject = JObject.Parse(json);
        JObject header = JObject.Parse(json);

        // Write data header
        ExcelWorksheet ws = package.Workbook.Worksheets["#HEADER#"];
        ws = package.Workbook.Worksheets.Add("#HEADER#");

        // Remove Application Header
        JObject hdr = (JObject)header.First.First.Last.First;
        hdr.Remove("Application_Header");

        // Write and Hide Header
        JObject headerJSON = new JObject();
        headerJSON["name"] = fileName;
        headerJSON["header"] = header;
        ws.Cells[1, 1].Value = headerJSON.ToString();
        ws.Hidden = eWorkSheetHidden.VeryHidden;

        // Start Recursive Looping with parameter Application Header as JObject
        converterV2.addSheet((JObject)jsonObject.First.First.Last.First, package, package.Workbook.Worksheets.Add("Request"), 1, "-", 0);
    }

    private dynamic BrowseButton(object sender, RoutedEventArgs e, string ext = "", bool allowedMultipleFiles = false)
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

        OpenFileDialog dlg = new OpenFileDialog { Filter = filter, Multiselect = allowedMultipleFiles };

        // Display OpenFileDialog by calling ShowDialog method 
        Nullable<bool> result = dlg.ShowDialog();

        // Arrange Path and File Name
        JObject file = new JObject();
        JArray files = new JArray();

        // If Allowed MultipleFiles
        if (result == true)
        {
            if (allowedMultipleFiles)
            {
                // Open document 
                foreach (string filePath in dlg.FileNames)
                {
                    JObject fileProperties = new JObject();
                    fileProperties["path"] = filePath;
                    fileProperties["name"] = filePath.Split("\\").Last().Split(".").First();
                    files.Add(fileProperties);
                }
            } else
            {
                file["path"] = dlg.FileName;
                file["name"] = dlg.FileName.Split("\\").Last().Split(".").First();
            }
        }

        return allowedMultipleFiles ? files : file;
    }

    private void btnSendRequestToAPI_Click(object sender, RoutedEventArgs e)
    {
        if (cb_endpoint.Text == "")
        {
            MessageBox.Show("[WARNING]: Please select an endpoint!");
        }
        else
        {
            // Flush response list item
            lb_responseList.Items.Clear();

            // Send Request to API
            List<Item> selectedRequestItem = lb_requestItems.FindAll(item => item.isSelected == true);
            int iterator = 0;
            if (selectedRequestItem.Count > 0)
            {
                foreach (Item it in selectedRequestItem)
                {
                    postRequestCRDE(it.json, it.fileName, iterator);
                }
            } else
            {
                MessageBox.Show("[WARNING]: Please select at least one request to send!");
            }
        }
    }
    private async void postRequestCRDE(string json, string saveFileNameResponse, int iterator)
    {
        saveFileNameResponse = saveFileNameResponse + "_response";
        // API endpoint
        string apiUrl = CRDE.ENDPOINT_REQUEST;

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
                converter.saveTextFile(@"\output\json\response\" + saveFileNameResponse + ".json", responseJsonIndent);

                // Convert Response to Excel
                convertJSONToExcel(package, responseJsonText, saveFileNameResponse, iterator);

                // Save Excel file
                string excelFilePath = GeneralMethod.getProjectDirectory() + @"\output\excel\response\" + saveFileNameResponse + '-' + GeneralMethod.getTimeStampNow() + ".xlsx";
                package.SaveAs(new FileInfo(excelFilePath));

                // Add to List Box Response
                lb_responseList.Items.Add(new Item { fileName = saveFileNameResponse, json = json, isSelected = false });

                MessageBox.Show("[SUCCESS]: [" + saveFileNameResponse + @"] Save Response was successful! File saved to \output\json\response and \output\excel\response");
            }
        }
        catch (HttpRequestException ex)
        {
            MessageBox.Show($"[API_FAILED]: {ex.StatusCode} : {ex.Message}", "Error");

        }
        catch (Exception ex)
        {
            MessageBox.Show($"[API_FAILED]: An error occurred: {ex.Message}", "Error");

        }
    }

    private void CheckBox_Click(object sender, RoutedEventArgs e)
    {
        foreach (Item item in lb_requestItems)
        {
            item.isSelected = (bool) cb_selectAll.IsChecked;
        }

        lb_requestList.Items.Refresh();
    }
}