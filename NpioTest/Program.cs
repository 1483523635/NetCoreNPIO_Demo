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

        public  static List<Goods> goodsList=new List<Goods>();
        // public static List<Goods> goodsList2=new List<Goods>();
        static void Main(string[] args)
        {
            Console.WriteLine("start");
          var list1=  File("1.xlsx");

           var list2=  File("2.xlsx");
            //foreach (var l in list1)
            //{
            //    Console.WriteLine(l.rowNum);
            //}
            //Console.WriteLine(list1.Count);
            //Console.WriteLine(list2.Count);
            Console.WriteLine(goodsList.Count);
            //Console.ReadLine();
            var goodsList_notContain = new List<Goods>();
            
            var goodList_total = GetTotalFile("3.xlsx");
            Console.WriteLine($" total is {goodList_total.Count}");
            Console.WriteLine($" distinct is {goodList_total.Distinct().Count()}");
            Console.Read();
            Console.WriteLine("read done");
            goodList_total.ForEach(g =>
            {
                if (!goodsList.Any(g1=>g1.codeNum==g.codeNum))
                goodsList_notContain.Add(g);
            });
            goodsList_notContain.ForEach(g => { Console.WriteLine(g.rowNum);});
            Console.WriteLine("total is "+goodsList_notContain.Count);
            Console.Read();
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
                        goodsList.Add(goods);
                    }
                }

            }
            return list;
        }

        static List<Goods> GetTotalFile(string filePath)
        {
            var goodsList = new List<Goods>();
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
                        if (row.FirstCellNum < 0) continue;
                        
                        if (!int.TryParse(row.GetCell(row.FirstCellNum).ToString(), out int id)) continue;
                        if (row.GetCell(row.FirstCellNum + 3) == null) continue;
                        goods.rowNum = id;
                        
                       
                        goods.codeNum = row.GetCell(row.FirstCellNum + 3).ToString();
                        goods.count = 0F;
                        goodsList.Add(goods);
                        if (i == 1263)
                        {
                            break;
                        }
                    }
                   
                }
            }
            return goodsList;
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
