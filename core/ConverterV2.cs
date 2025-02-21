using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRDEConverterJsonExcel.core
{
    class ConverterV2
    {
        public void addSheet(JObject data, ExcelPackage package, ExcelWorksheet worksheet = null, int startRow = 1, string parent = "", int parentId = 1, string parentName = "", int levelNested = 0)
        {
            foreach (var property in data)
            {
                if (property.Key == "Variables")
                {
                    if (property.Value.Count() == 0)
                    {
                        worksheet.Cells[startRow, 1].Value = startRow;
                        worksheet.Cells[startRow, 2].Value = parent;
                        worksheet.Cells[startRow, 3].Value = parentId;
                    }

                    // Write The Cells
                    foreach (var variable in (JObject)property.Value)
                    {
                        worksheet.Cells[startRow, 4].Value = variable.Value.Type.ToString();
                        worksheet.Cells[startRow, 5].Value = variable.Key;
                        worksheet.Cells[startRow, 6].Value = variable.Value.ToString();
                        startRow++;
                    }
                }
                else if (property.Key == "Categories")
                {
                    foreach (var category in property.Value)
                        addSheet((JObject)category, package, worksheet, startRow + 1, parentName, parentId, "", ++levelNested);
                }
                else
                {
                    // Print Header
                    worksheet.Cells[startRow, 1].Value = startRow;
                    worksheet.Cells[startRow, 2].Value = parent;
                    worksheet.Cells[startRow, 3].Value = parentId;
                    worksheet.Cells[startRow, 4].Value = property.Key;

                    // Coloring Header Backround Cell
                    for(int col = 1; col <= 6; col++)
                    {
                        worksheet.Cells[startRow, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[startRow, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.SkyBlue);
                    }
                    //Trace.WriteLine("Start Row : " + package.Workbook.Worksheets["Request"].Dimension.End.Row + 1 + "||Header: " + property.Key);

                    //int endRow = package.Workbook.Worksheets[property.Key].Dimension == null ? 1 : package.Workbook.Worksheets["Request"].Dimension.End.Row;
                    addSheet((JObject)property.Value, package, worksheet, package.Workbook.Worksheets["Request"].Dimension.End.Row + 1, parent, package.Workbook.Worksheets["Request"].Dimension.End.Row, property.Key, levelNested);
                }
            }
        }

        public JArray normalizeDataFromJSON(JArray result, JObject jsonObject, int id, string parent, int parentId)
        {
            foreach (var property in jsonObject)
            {
                if (property.Key == "Variables")
                {
                    JObject normalizeVariableData = new JObject();

                    normalizeVariableData["Id"] = id;
                    normalizeVariableData["Parent"] = parent;
                    normalizeVariableData["ParentId"] = parentId;

                    foreach (var variable in (JObject)property.Value)
                        normalizeVariableData[variable.Key] = variable.Value;

                    result.Add(normalizeVariableData);
                }
                else if (property.Key == "Categories")
                {
                    //foreach (var category in property.Value)
                    //    normalizeDataFromJSON(result, (JObject)category);
                } 
                else
                {
                    normalizeDataFromJSON(result, (JObject)property.Value, result.Count + 1, property.Key, id);
                }
            }

            return result;
        }
    }
}
