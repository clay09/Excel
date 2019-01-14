using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelChange
{
    public class SupplierModels
    {
        /// <summary> 資料 </summary>
        public class BuyerInfoModel
        {
            public string Name { get; set; }
            public string Season { get; set; }
            public string Year   { get; set; }
            public string Date { get; set; }

            public BuyerInfoModel GetInfo(string path)
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

        /// <summary> HNG 供應商 </summary>
        public class HNGModel
        {
            public string Supplier { get; set; }
            public string PDM { get; set; }
            public string Description { get; set; }
            public string Unit { get; set; }
            public string Color { get; set; }
            public string Style { get; set; }
            public string DIM { get; set; }
            public string Size { get; set; }
            public string Zipper { get; set; }
            public string Quantity { get; set; }
            public string Unitprice { get; set; }
            public string Amount { get; set; }
        }
    }
}
