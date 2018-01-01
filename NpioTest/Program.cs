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

        public static List<Goods> goodsList = new List<Goods>();
        // public static List<Goods> goodsList2=new List<Goods>();
        static void Main(string[] args)
        {
            Console.WriteLine("start");
            var list1 = File("1.xlsx");
            var list2 = File("2.xlsx");
          
            var goodsList_notContain = new List<Goods>();
            var goodList_total = GetTotalFile("3.xlsx");
            Console.WriteLine("read done");
            goodList_total.ForEach(g =>
            {
                if (!goodsList.Any(g1 => g1.codeNum == g.codeNum))
                    goodsList_notContain.Add(g);
            });
           
            ExportToExcel(goodsList_notContain);
            Console.WriteLine("success!");
            Console.WriteLine("total is " + goodsList_notContain.Count);
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
                    for (int i = 0; i <= cellCount; ++i)
                    {

                        var goods = new Goods();
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　
                        if (row.FirstCellNum < 0) continue;

                        if (!int.TryParse(row.GetCell(row.FirstCellNum).ToString(), out int id)) continue;
                        if (row.GetCell(row.FirstCellNum + 3) == null) continue;
                        goods.rowNum = i;

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


        static void ExportToExcel(List<Goods> data)
        {
            using (var fs_write = new FileStream("tatal.xls", FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                IWorkbook workbook_write = new HSSFWorkbook();
                ISheet sheet_write = workbook_write.CreateSheet("toal");

                using (var fs = new FileStream("3.xlsx", FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new HSSFWorkbook(fs);
                    int i = 0;
                    var sheet = workbook.GetSheetAt(0);
                    var row_header = sheet.GetRow(2);
                    var row_header_write = sheet_write.CreateRow(i);
                    for (int j = 0; j < row_header.LastCellNum; j++)
                    {
                        row_header_write.CreateCell(j).SetCellValue(row_header.GetCell(j).ToString());
                    }
                    data.ForEach(g =>
                    {
                         i++;
                        var row_read = sheet.GetRow(g.rowNum);
                        var row_write = sheet_write.CreateRow(i);
                        for (int j = 0; j < row_read.LastCellNum ; j++)
                        {
                            row_write.CreateCell(j).SetCellValue(row_read.GetCell(j).ToString());
                        }
                        
                       
                    });
                }
                workbook_write.Write(fs_write);
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
