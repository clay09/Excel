using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using static ExcelChange.SupplierModels;

namespace ExcelChange
{
    class Program
    {
        


        static void Main(string[] args)
        {
            var path = @"C:\Users\ClayChen\Desktop\Grace\HNG-F19-1220 ~原始單.xls";
            GetBuyerName(path);
            var excel = GetExcel(path);
            var data = ReadFromExcelFile(excel);

            WriteToExcel(@"C:\Users\ClayChen\Desktop\NewExcel\HNG-F19.xls", data);
        }

        /// <summary>
        /// 讀取excel文檔
        /// </summary>
        /// <param name="data">文檔路徑</param>
        public static List<HNGModel> ReadFromExcelFile(IWorkbook data)
        {
            List<HNGModel> result = null;
            try
            {
                var sheetCount = data.NumberOfSheets; // 取的所有頁籤數量
                for (int g = 0; g < 1/*sheetCount*/; g++)
                {
                    ISheet sheet = data.GetSheetAt(g); //讀取當前頁籤(表)數據

                    result = GetSingleSheetData(sheet);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return result;
        }

        /// <summary> 取得當前頁籤資料 </summary>
        /// <param name="data">當前頁籤資料</param>
        public static List<HNGModel> GetSingleSheetData(ISheet data)
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

           return DataTransfer(rowStringValue);
        }


        public static List<HNGModel> DataTransfer(List<string> data)
        {
            List<HNGModel> supplierData = new List<HNGModel>();
            try
            {
                string temp = null;
                for (int i = 1; i < data.Count; i++)
                {
                    
                    var itemArray = data[i].Split('\r');
                    var HNG = new HNGModel();
                    HNG.Area = "1";
                    HNG.Supplier = itemArray[0];
                    HNG.PDM_SerialNO = "1";
                    HNG.PDM_NO = itemArray[1];
                    HNG.Style = itemArray[5];
                    HNG.Description = "";
                    HNG.YKK_clr_code = "";
                    HNG.NIKE_clr_code = "";
                    if (itemArray[4] == "沒資料" && temp == null)
                    {
                        temp = data[i - 1].Split('\r')[4];
                        HNG.Color_Description = temp;
                    }
                    else if (itemArray[4] == "沒資料" && temp != null)
                    {
                        HNG.Color_Description = temp;
                    }
                    else
                    {
                        HNG.Color_Description = itemArray[4];
                        temp = null;
                    }
                    HNG.Size = itemArray[7];
                    HNG.Length = "";
                    HNG.Length_Unit = "";
                    HNG.Ins = "";
                    HNG.Gender = "MAN";
                    HNG.Qty = itemArray[9];
                    HNG.Unit = itemArray[3];
                    HNG.QRS_PP_Qty = "";
                    HNG.CC_Qty = "";
                    HNG.Sample_Qty = "";
                    HNG.U_Price = "";
                    HNG.Sp_UnitPrice = "";
                    HNG.Amount = "";
                    HNG.YKKItemCode = "";
                    HNG.NIKE_NO = "";
                    HNG.NIKE_Meterial = "";
                    supplierData.Add(HNG);
                }
            }
            catch (Exception ex)
            {
                throw;
            }

            return supplierData;
        }


        public static void WriteToExcel(string filePath, List<HNGModel> data)
        {
            //取得excel 
            IWorkbook wb = GetExcel(filePath);
            
            //創建一個表單
            ISheet sheet = wb.CreateSheet("Sheet7");
            //設置列寬
            int[] columnWidth = { 10, 10, 10, 10, 10, 20, 10, 10, 10, 10, 10, 10, 20, 10, 10, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2 };
            for (int i = 0; i < columnWidth.Length; i++)
            {
                //設置列寬度，256*字符數，因為單位是1/256個字符
                sheet.SetColumnWidth(i, 256 * columnWidth[i]);
            }
            
            // 表頭
            Header(sheet);

            // 報表內容
            Body(data, sheet);

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

        public static void Body(List<HNGModel> data, ISheet sheet)
        {
            Type modelType = new HNGModel().GetType();
           
            int count = 1;
            
            foreach (var item in data)
            {
                var jh = modelType.GetProperties().Length;
                for (int i = 0; i < modelType.GetProperties().Length; i++)
                {
                    var propValue = modelType.GetProperties()[i].GetValue(item, null)?.ToString();
                    sheet.CreateRow(count).CreateCell(i).SetCellValue(propValue);
                }
                count++;
            }

        }

        public static void Header(ISheet sheet)
        {
            IRow row;
            ICell cell;
            Type myType = typeof(HNGModel);
            int j = 0;

            // 表頭
            for (int i = 0; i < 1; i++)
            {
                row = sheet.CreateRow(i);//創建第i行
                foreach (var item in myType.GetProperties())
                {
                    cell = row.CreateCell(j);//創建第j列
                    cell.SetCellValue(item.Name);
                    j++;
                }
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

        /// <summary>
        /// 根據指定的文檔格式創建對應的類
        /// </summary>
        /// <param name="extension">文檔路徑</param>
        /// <returns></returns>
        public static IWorkbook ExcelType(string extension)
        {
            //創建工作薄  
            IWorkbook wb;

            //根據指定的文檔格式創建對應的類
            if (extension.Equals(".xls"))
            {
                wb = new HSSFWorkbook();
            }
            else
            {
                wb = new XSSFWorkbook();
            }

            return wb;
        }


        public void g()
        {
            //ICellStyle style1 = wb.CreateCellStyle();//樣式
            //style1.Alignment = HorizontalAlignment.Left;//文本水平對齊方式
            //style1.VerticalAlignment = VerticalAlignment.Center;//文本垂直對齊方式
            //                                                    //設置邊框
            //style1.BorderBottom = BorderStyle.Thin;
            //style1.BorderLeft = BorderStyle.Thin;
            //style1.BorderRight = BorderStyle.Thin;
            //style1.BorderTop = BorderStyle.Thin;
            //style1.WrapText = true;//自動換行
            //ICellStyle style2 = wb.CreateCellStyle();//樣式
            //IFont font1 = wb.CreateFont();//字體
            //font1.FontName = "楷體";
            //font1.Color = HSSFColor.Red.Index;//字體顏色
            //font1.Boldweight = (short)FontBoldWeight.Normal;//字體加粗樣式
            //style2.SetFont(font1);//樣式裏的字體設置具體的字體樣式
            //                      //設置背景色
            //style2.FillForegroundColor = HSSFColor.Yellow.Index;
            //style2.FillPattern = FillPattern.SolidForeground;
            //style2.FillBackgroundColor = HSSFColor.Yellow.Index;
            //style2.Alignment = HorizontalAlignment.Left;//文本水平對齊方式
            //style2.VerticalAlignment = VerticalAlignment.Center;//文本垂直對齊方式
            //ICellStyle dateStyle = wb.CreateCellStyle();//樣式
            //dateStyle.Alignment = HorizontalAlignment.Left;//文本水平對齊方式
            //dateStyle.VerticalAlignment = VerticalAlignment.Center;//文本垂直對齊方式
            //                                                       //設置數據顯示格式
            //IDataFormat dataFormatCustom = wb.CreateDataFormat();
            //dateStyle.DataFormat = dataFormatCustom.GetFormat("yyyy-MM-dd HH:mm:ss");
        }
    }
}
