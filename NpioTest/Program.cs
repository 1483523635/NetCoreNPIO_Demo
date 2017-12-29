using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NpioTest
{
    class Program
    {

        // public  static List<Goods> goodsList=new List<Goods>();
        // public static List<Goods> goodsList2=new List<Goods>();
        static void Main(string[] args)
        {
            Console.WriteLine("start");
            var goodsList = File("1.xlsx");

            var goodsList2 = File("2.xlsx");
            Console.WriteLine("read done");
            foreach (var g in goodsList2)
            {
                var good_list = new List<Goods>();
                good_list = goodsList.Where(good => good.codeNum == g.codeNum).ToList();
                if (good_list == null || good_list.Count <= 0)
                {
                    Console.WriteLine($"table 2， row num is {g.rowNum} ,not contain in table 1");
                    continue;
                }
                float sum = 0F;
                foreach (var g3 in good_list)
                {
                    sum += g3.count;
                }
                good_list.ForEach(g1 => sum += g1.count);
                if (sum != g.count)
                {
                    foreach (var g2 in good_list)
                    {
                        Console.WriteLine($"table 2 ,row num is {g.rowNum}, not match in table 1 row num is {g2.rowNum} ,place check");
                    }

                }
            }
            Console.Read();
            Console.WriteLine("done！");

            Console.Read();
            Console.WriteLine("Hello World!");
        }

        static List<Goods> File(string filePath)
        {
            var list = new List<Goods>();
            //var filePath = "1.xlsx";
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
            {

                IWorkbook workbook = new HSSFWorkbook(fs);

                for (int k = 0; k < workbook.NumberOfSheets; k++)
                {
                    var sheet = workbook.GetSheetAt(k);
                    int cellCount = sheet.LastRowNum;
                    for (int i = 1; i <= cellCount; ++i)
                    {
                        var goods = new Goods();
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　
                        if (row.FirstCellNum <= 0) continue;

                        if (!int.TryParse(row.GetCell(row.FirstCellNum).ToString(), out int id)) continue;
                        goods.rowNum = id;
                        goods.codeNum = row.GetCell(row.FirstCellNum + 1).ToString();
                        if (float.TryParse(row.GetCell(row.FirstCellNum + 5).ToString(), out float count))
                        {
                            goods.count = count;
                        }

                        list.Add(goods);
                    }
                }

            }
            return list;
        }

        static void GetTotalFile(string filePath)
        {
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
            {
                IWorkbook workbook = new HSSFWorkbook(fs);
                var goods = new Goods();
                for (int k = 0; k < workbook.NumberOfSheets; k++)
                {
                    var sheet = workbook.GetSheetAt(k);
                    int cellCount = sheet.LastRowNum;
                    for (int i = 1; i <= cellCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　
                        if (row.FirstCellNum <= 0) continue;

                        if (!int.TryParse(row.GetCell(row.FirstCellNum).ToString(), out int id)) continue;
                        goods.rowNum = id;
                        goods.codeNum = row.GetCell(row.FirstCellNum + 3).ToString();
                        if (float.TryParse(row.GetCell(row.FirstCellNum + 5).ToString(), out float count))
                        {
                            goods.count = count;
                        }

                        //goodsTotalList.Add(goods);
                    }
                }

            }
        }
    }


    /// <summary>
    /// 商品类
    /// </summary>
    public class Goods
    {
        /// <summary>
        /// 行号
        /// </summary>
        public int rowNum { get; set; }
        /// <summary>
        /// 编号
        /// </summary>
        public string codeNum { get; set; }
        /// <summary>
        /// 数量
        /// </summary>
        public float count { get; set; }
    }
}
