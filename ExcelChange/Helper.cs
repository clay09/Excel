using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using static ExcelChange.SupplierModels;

namespace ExcelChange
{
    public class Helper
    {
        static IRow row;
        static ICell cell;

        /// <summary> 讀取excel文檔，根據指定的文檔格式創建對應的類 </summary>
        /// <param name="filePath">檔案路徑</param>
        /// <returns></returns>
        public static IWorkbook GetExcel(string filePath)
        {
            IWorkbook result = null;
            string extension = Path.GetExtension(filePath);

            FileStream fs = File.OpenRead(filePath);

            if (extension.Equals(".xls")) // 2003版本
            {
                result = new HSSFWorkbook(fs); // 把xls文檔中的數據寫入wk中
            }
            else // 2007版本
            {
                result = new XSSFWorkbook(fs); // 把xlsx文檔中的數據寫入wk中
            }

            fs.Close();

            return result;
        }

        /// <summary>
        /// 表頭
        /// </summary>
        /// <param name="sheet"></param>
        public static void Header<T>(ISheet sheet, int rowIndex, T pObject)
        {
            Type myType = pObject.GetType();
            int j = 0;

            row = sheet.CreateRow(rowIndex);//創建第1行

            foreach (var item in myType.GetProperties())
            {
                cell = row.CreateCell(j);//創建第j列
                cell.SetCellValue(item.Name);
                j++;
            }
        }

        /// <summary>
        /// 資料本體
        /// </summary>
        /// <param name="data"></param>
        /// <param name="sheet"></param>
        public static void Body<T>(List<T> data, ISheet sheet, T pObject)
        {
            Type modelType = pObject.GetType();
            int count = 1;

            foreach (var item in data)
            {
                row = sheet.CreateRow(count); // 創建第x行
                for (int i = 0; i < modelType.GetProperties().Length; i++)
                {
                    var propValue = modelType.GetProperties()[i].GetValue(item, null)?.ToString();

                    cell = row.CreateCell(i); // 創建第i列
                    cell.SetCellValue(propValue);
                    sheet.AutoSizeColumn(i); // 如果要根據內容自動調整列寬，需要先setCellValue再調用
                }
                count++;
            }
        }
    }
}
