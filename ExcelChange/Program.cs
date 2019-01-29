using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static ExcelChange.SupplierModels;

namespace ExcelChange
{
    class Program
    {
        static void Main(string[] args)
        {
            var path = @"C:\Users\ClayChen\Desktop\Grace\HNG-F19-1220 ~原始單.xls";
            GetBuyerName(path);
            var excel = Helper.GetExcel(path);
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
                for (int i = 10; i < 11/*sheetCount*/; i++)
                {
                    ISheet sheet = data.GetSheetAt(i); //讀取當前頁籤(表)數據

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
                    break;
                
                row = data.GetRow(i);  //讀取當前行數據，LastRowNum 是當前表的總行數-1（注意）

                if (row != null)
                {
                    #region 取得資料
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
                    #endregion

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
                string temp = null; string color = null; string gender = null;
                for (int i = 1; i < data.Count; i++)
                {
                    var itemArray = data[i].Split('\r');
                    color = Color(itemArray[4]);
                    if (gender == null)
                    {
                        gender = Gender(itemArray[4]);
                    }

                    var HNG = new HNGModel();
                    HNG.Area = "1";
                    HNG.Supplier = itemArray[0];
                    HNG.PDM_SerialNO = null;
                    HNG.PDM_NO = itemArray[1];
                    HNG.Style = itemArray[5];
                    #region MyRegion
                    //if (itemArray[4] == "沒資料" && temp == null)
                    //{
                    //    temp = data[i - 1].Split('\r')[4];
                    //    HNG.Color_Description = temp;
                    //}
                    //else if (itemArray[4] == "沒資料" && temp != null)
                    //{
                    //    HNG.Color_Description = temp;
                    //}
                    //else
                    //{
                    //    HNG.Color_Description = itemArray[4];
                    //    temp = null;
                    //}
                    #endregion
                    if (color == "沒資料" && temp == null)
                    {
                        temp = Color(data[i - 1].Split('\r')[4]);
                        HNG.Color_Description = temp;
                    }
                    else if (color == "沒資料" && temp != null)
                    {
                        HNG.Color_Description = temp;
                    }
                    else
                    {
                        HNG.Color_Description = color;
                        temp = null;
                    }
                    HNG.Size = itemArray[7] == "沒資料" ? string.Empty : itemArray[7];
                  
                    HNG.Qty = itemArray[9] == "沒資料" ? string.Empty : itemArray[9];
                    HNG.Unit = itemArray[3];
                   
                    supplierData.Add(HNG);
                }

                var groupbyData = supplierData.GroupBy(a => new { a.Supplier, a.PDM_NO });

                int no = 1;
                foreach (var item in groupbyData)
                {
                    foreach (var groupItem in item)
                    {
                        groupItem.PDM_SerialNO = no.ToString();
                        groupItem.Gender = gender == null ? string.Empty : gender;
                        no++;
                    }
                    no = 1;
                }
            }
            catch (Exception ex)
            {
                throw;
            }

            return supplierData.OrderBy(a => a.Supplier).ToList(); 
        }

        public static string Do(List<string> allData, string[] data, int index)
        {
            string result;
            string temp = null;
            if (data[index] == "沒資料" && temp == null)
            {
                temp = allData[index - 1].Split('\r')[index];
                result = temp;
            }
            else if (data[index] == "沒資料" && temp != null)
            {
                result = temp;
            }
            else
            {
                result = data[index];
                temp = null;
            }

            return result;
        }

        public static void WriteToExcel(string filePath, List<HNGModel> data)
        {
            //取得excel 
            IWorkbook wb = Helper.GetExcel(filePath);
            
            //創建一個表單
            ISheet sheet = wb.CreateSheet("Sheet1");
           
            // 表頭
            Helper.Header(sheet, 0, new HNGModel());

            // 報表內容
            Helper.Body(data, sheet, new HNGModel());

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
        /// 取得顏色
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static string Color(string data)
        {
            // 檢查有沒有gender的資料在裡面，有的話移除掉gender的資料
            string result = data?.ToLower().IndexOf("gender") > -1 ?
                               data.Substring(0, data.ToLower().IndexOf("gender") - 1) :
                               data;

            return result;
        }

        /// <summary>
        /// 取得性別
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static string Gender(string data)
        {
            string result = null;
            if (data?.ToLower().IndexOf("gender") > -1)
            {
                result = data
                    .Substring(data.ToLower().IndexOf("gender"))
                    .ToLower().Contains("w") ?
                    "Women" : "Men";
            }

            return result;
        }
    }
}
