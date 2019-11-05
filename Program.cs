using System;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using LitJson;
using System.Text;

namespace ExcelToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please at least input one file!");
                return;
            }
            for (int i = 0; i < args.Length; ++i)
            {
                string fileName = args[i];
                if (fileName.EndsWith("xlsx") || fileName.EndsWith("xls"))
                {
                    if (File.Exists(fileName))
                    {
                        FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                        XSSFWorkbook workbook = new XSSFWorkbook(file);
                        WriteJson(workbook);
                    }
                    else
                    {
                        Console.WriteLine("{0:G} not exit!", fileName);
                    }
                }
                else
                {
                    Console.WriteLine("file must be end with xlsx or xls!");
                }
            }
            return;
        }


        static void WriteJson(XSSFWorkbook workbook)
        {
            //如果导出目录不存在，创建目录
            string outFolder = @"json\";
            if (!Directory.Exists(outFolder))
            {
                Directory.CreateDirectory(outFolder);
            }

            //服务端表头位置
            var server_title = 1;
            //服务端数据开始行数
            var server_num = 2;

            //获取excel的sheet
            for (int m = 0; m < workbook.NumberOfSheets; ++m)
            {
                string sheet_name = workbook.GetSheetName(m);

                ISheet sheet = workbook.GetSheetAt(m);
                try
                {
                    string txtPath = @"json\" + sheet_name + ".json";
                    FileStream aFile = new FileStream(txtPath, FileMode.Create);
                    StreamWriter sw = new StreamWriter(aFile,Encoding.UTF8); //直接保存utf8文件
                    JsonData arrData = new JsonData();

                    //获取sheet的第二行，服务端用的表头
                    IRow titleRow = sheet.GetRow(server_title);

                    //一行最后一个方格的编号 即总的列数
                    int cellCount = titleRow.LastCellNum;

                    //最后一行
                    int rowCount = sheet.LastRowNum;

                    //遍历行
                    for (int i = server_num; i <= rowCount; i++)
                    {
                        IRow row = sheet.GetRow(i);

                        if (row == null)
                        {
                            Console.WriteLine("{0:G} row {1:D} have exception data!", sheet_name, i);
                            break;
                        }
                        JsonData rowData = new JsonData();
                        arrData.Add(rowData);
                        //遍历该行的列
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            if (titleRow.GetCell(j) != null && titleRow.GetCell(j).ToString().Length != 0)
                            {
                                if (row.GetCell(j) != null)
                                {
                                    if (row.GetCell(j).CellType == CellType.String)
                                    {
                                        string value = row.GetCell(j).StringCellValue;
                                        
                                        string title = titleRow.GetCell(j).ToString().Trim();
                                        //value = value.Replace("\r", "").Replace("\n", "\\n");
                                        rowData[title] = value;
                                    }
                                    else if (row.GetCell(j).CellType == CellType.Numeric)
                                    {
                                        double value = row.GetCell(j).NumericCellValue;
                                        string title = titleRow.GetCell(j).ToString().Trim();
                                        rowData[title] = value;
                                    }
                                }
                            }
                        }
                    }
                    string jsonStr = arrData.ToJson();
                    sw.Write(jsonStr);
                    sw.Close();
                    Console.WriteLine("{0:G} success!", sheet_name);
                }
                catch (IOException ex)
                {
                    Console.WriteLine("something error!");
                    return;
                }
            }
        }
    }
}
