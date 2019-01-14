using ExcelChange;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ExcelChange.SupplierModels;

namespace ExcelChange
{
    class Program
    {
        static void Main(string[] args)
        {
            var path = @"C:\Users\User\Desktop\Grace\HNG-F19-1220 ~原始單.xls";
            //var path = @"C:\Users\ClayChen\Desktop\測試123.xlsx";
            GetBuyerName(path);
            var data = GetExcel(path);
            ReadFromExcelFile(data);
        }

        /// <summary>
        /// 讀取excel文檔
        /// </summary>
        /// <param name="data">文檔路徑</param>
        public static void ReadFromExcelFile(IWorkbook data)
        {
            try
            {
                var sheetCount = data.NumberOfSheets; // 取的所有頁籤數量
                for (int g = 0; g < 1/*sheetCount*/; g++)
                {
                    ISheet sheet = data.GetSheetAt(g); //讀取當前頁籤(表)數據

                    GetSingleSheetData(sheet);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary> 取得當前頁籤資料 </summary>
        /// <param name="data">當前頁籤資料</param>
        public static void GetSingleSheetData(ISheet data)
        {
            IRow row;

            string rowText = string.Empty; // 當前(行)的資料
            string cellText = string.Empty; // 當前儲存格的資料 
            List<string> rowStringValue = new List<string>();

            bool start = false;
            bool isFinish = false;
            for (int i = 25; i <= data.LastRowNum; i++) // 處理橫的(行)
            {
                if (isFinish)
                {
                    break;
                }

                row = data.GetRow(i);  //讀取當前行數據，LastRowNum 是當前表的總行數-1（注意）

                if (row != null)
                {
                    for (int j = 0; j < row.LastCellNum; j++) //處理(欄位)，LastCellNum 是當前行的總列數 
                    {

                        row.GetCell(j)?.SetCellType(CellType.String); // 設定欄位資料型態為字串
                        cellText = row.GetCell(j)?.StringCellValue.Replace("\r\n", "");
                        //讀取該行的第j列數據
                        string value = String.IsNullOrEmpty(cellText) ? "沒資料" : cellText;

                        // 判斷是否到了表格的標題欄位，是的話就跳出，準備收集下一行的正式資料
                        if (value == "SUPPLIER")
                        {
                            start = true;
                            break;
                        }
                        // 判斷是否到表格結尾，是的話就跳出，結束迴圈
                        if (value == "TOTAL")
                        {
                            isFinish = true; // 結束迴圈迴圈
                            start = false; // 結束收集資料
                            break;
                        }
                        // 條件成立就開始收集資料
                        if (start)
                        {
                            rowText += value + "\r";
                        }
                    }

                    // 條件成立就將收集的資料加入集合中
                    if (start)
                    {
                        rowStringValue.Add(rowText);
                        rowText = string.Empty;
                    }
                }
            }

            DataTransfer(rowStringValue);
        }


        public static void DataTransfer(List<string> data)
        {
            try
            {
                List<HNGModel> supplierData = new List<HNGModel>();
                string temp = null;
                for (int i = 1; i < data.Count; i++)
                {
                    var itemArray = data[i].Split('\r');
                    var gy = new HNGModel();
                    gy.Supplier = itemArray[0];
                    gy.PDM = itemArray[1];
                    gy.Description = itemArray[2];
                    gy.Unit = itemArray[3];

                    if (itemArray[4] == "沒資料" && temp == null)
                    {
                        temp = data[i - 1].Split('\r')[4];
                        gy.Color = temp;
                    }
                    else if (itemArray[4] == "沒資料" && temp != null)
                    {
                        gy.Color = temp;
                    }
                    else
                    {
                        gy.Color = itemArray[4];
                        temp = null;
                    }

                    gy.Style = itemArray[5];
                    gy.DIM = itemArray[6];
                    gy.Size = itemArray[7];
                    gy.Zipper = itemArray[8];
                    gy.Quantity = itemArray[9];
                    gy.Unitprice = itemArray[10];
                    gy.Amount = itemArray[11];

                    supplierData.Add(gy);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }


        public void WriteToExcel(string filePath)
        {
            //創建工作薄  
            IWorkbook wb;
            string extension = System.IO.Path.GetExtension(filePath);
            //根據指定的文檔格式創建對應的類
            if (extension.Equals(".xls"))
            {
                wb = new HSSFWorkbook();
            }
            else
            {
                wb = new XSSFWorkbook();
            }
            ICellStyle style1 = wb.CreateCellStyle();//樣式
            style1.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;//文本水平對齊方式
            style1.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文本垂直對齊方式
                                                                                  //設置邊框
            style1.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style1.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style1.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style1.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            style1.WrapText = true;//自動換行
            ICellStyle style2 = wb.CreateCellStyle();//樣式
            IFont font1 = wb.CreateFont();//字體
            font1.FontName = "楷體";
            font1.Color = HSSFColor.Red.Index;//字體顏色
            font1.Boldweight = (short)FontBoldWeight.Normal;//字體加粗樣式
            style2.SetFont(font1);//樣式裏的字體設置具體的字體樣式
                                  //設置背景色
            style2.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            style2.FillPattern = FillPattern.SolidForeground;
            style2.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            style2.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;//文本水平對齊方式
            style2.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文本垂直對齊方式
            ICellStyle dateStyle = wb.CreateCellStyle();//樣式
            dateStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;//文本水平對齊方式
            dateStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文本垂直對齊方式
                                                                                     //設置數據顯示格式
            IDataFormat dataFormatCustom = wb.CreateDataFormat();
            dateStyle.DataFormat = dataFormatCustom.GetFormat("yyyy-MM-dd HH:mm:ss");
            //創建一個表單
            ISheet sheet = wb.CreateSheet("Sheet0");
            //設置列寬
            int[] columnWidth = { 10, 10, 20, 10 };
            for (int i = 0; i < columnWidth.Length; i++)
            {
                //設置列寬度，256*字符數，因為單位是1/256個字符
                sheet.SetColumnWidth(i, 256 * columnWidth[i]);
            }
            //測試數據
            int rowCount = 3, columnCount = 4;
            object[,] data = {
                                 {"列0", "列1", "列2", "列3"},
                                 {"", 400, 5.2, 6.01},
                                 {"", true, "2014-07-02", DateTime.Now}
                                 //日期可以直接傳字符串，NPOI會自動識別
                                 //如果是DateTime類型，則要設置CellStyle.DataFormat，否則會顯示為數字
                             };
            IRow row;
            ICell cell;

            for (int i = 0; i < rowCount; i++)
            {
                row = sheet.CreateRow(i);//創建第i行
                for (int j = 0; j < columnCount; j++)
                {
                    cell = row.CreateCell(j);//創建第j列
                    cell.CellStyle = j % 2 == 0 ? style1 : style2;
                    //根據數據類型設置不同類型的cell
                    object obj = data[i, j];
                    //SetCellValue(cell, data[i, j]);
                    //如果是日期，則設置日期顯示的格式
                    if (obj.GetType() == typeof(DateTime))
                    {
                        cell.CellStyle = dateStyle;
                    }
                    //如果要根據內容自動調整列寬，需要先setCellValue再調用
                    //sheet.AutoSizeColumn(j);
                }
            }
            //合併單元格，如果要合併的單元格中都有數據，只會保留左上角的
            //CellRangeAddress(0, 2, 0, 0)，合併0-2行，0-0列的單元格
            CellRangeAddress region = new CellRangeAddress(0, 2, 0, 0);
            sheet.AddMergedRegion(region);
            try
            {
                FileStream fs = File.OpenWrite(filePath);
                wb.Write(fs);//向打開的這個Excel文檔中寫入表單並保存。  
                fs.Close();
            }
            catch (Exception e)
            {
                
            }
        }

        /// <summary> 讀取excel文檔 </summary>
        /// <param name="filePath">檔案路徑</param>
        /// <returns></returns>
        public static IWorkbook GetExcel(string filePath)
        {
            IWorkbook result = null;
            string extension = System.IO.Path.GetExtension(filePath);

            FileStream fs = File.OpenRead(filePath);

            if (extension.Equals(".xls")) // 2003版本
            {
                result = new HSSFWorkbook(fs); //把xls文檔中的數據寫入wk中
            }
            else // 2007版本
            {
                result = new XSSFWorkbook(fs); //把xlsx文檔中的數據寫入wk中
            }

            fs.Close();

            return result;
        }

        /// <summary> 從檔名取得 廠商名稱、記節、日期 </summary>
        /// <param name="path">檔案路徑</param>
        public static BuyerInfoModel GetBuyerName(string path)
        {
            BuyerInfoModel result = new BuyerInfoModel();

            var pathArray = path.Split('\\');
            var detailArray = pathArray[pathArray.Length - 1].Split('-');

            result.Name = detailArray[0];
            result.Season = detailArray[1].Substring(0).ToUpper() == "S" ? "Spring" : "Fall";
            result.Year = $"20{detailArray[1].Substring(1, 2)}";
            result.Date = detailArray[2].Substring(0, 4);

            return result;
        }


    }
}
