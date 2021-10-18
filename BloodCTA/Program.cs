using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.CompilerServices;
using BigInteger = System.Numerics.BigInteger;
using System.Numerics;
using System.Runtime;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using Complex = System.Numerics.Complex;

namespace BloodCTA
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = Directory.GetCurrentDirectory();
            ofd.Filter = "テキストファイル (*.xlsx)|*.xlsx|"
                + "すべてのファイル (*.*)|*.*";
            ofd.FilterIndex = 2;
            ofd.Multiselect = false;
            Console.WriteLine("※ファイル要件");
            Console.WriteLine("① エクセルファイルの最終シートを読み込みます");
            Console.WriteLine("② A列：PID、B列：採血受付日時(時刻1)、C列：患者認証(時刻2)");
            Console.WriteLine("");

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                Console.WriteLine("※解析条件");
                Console.WriteLine("① 採血受付日時(時刻1)と患者認証(時刻2)の値が日時である事");
                Console.WriteLine("② 採血受付日時(時刻1)と患者認証(時刻2)が同日である事");
                Console.WriteLine("③ 患者認証(時刻2)が7時～17時の間");
                Console.WriteLine("④ 同日中にPIDが重複している場合は、時間が早い方のみ選択");
                Console.WriteLine("");

                try
                {
                    List<RowData> rowDatas = new List<RowData>();
                    Console.CursorVisible = false;
                    char[] bars = { '／', '―', '＼', '｜' };
                    Console.WriteLine(ofd.FileName);
                    Console.Write(bars[0]);
                    Console.Write("読み込み中..." + "{0, 4:d0}%", 0);
                    Console.SetCursorPosition(0, Console.CursorTop);
                    using (XLWorkbook workbook = new XLWorkbook(ofd.FileName))
                    {
                        IXLWorksheet worksheet = workbook.Worksheets.Last();
                        int lastRow = worksheet.LastRowUsed().RowNumber();
          
                        for (int i = 2; i <= lastRow; i++)
                        {
                            Console.Write(bars[i % 4]);
                            Console.Write("解析中......." + "{0, 4:d0}%", 100 * (i + 1) / lastRow);
                            Console.SetCursorPosition(0, Console.CursorTop);

                            IXLCell cell = worksheet.Cell(i, 1);
                            double pid = (double)cell.Value;
                            DateTime dt1;
                            DateTime dt2;
                            if (!DateTime.TryParse(worksheet.Cell(i, 2).Value.ToString(), out dt1)) continue;
                            if (!DateTime.TryParse(worksheet.Cell(i, 3).Value.ToString(), out dt2)) continue;
                            rowDatas.Add(new RowData() { pID = pid, dateTime1 = dt1, dateTime2 = dt2 });

                            cell = null;

                        }
                        worksheet = null;
                        workbook.Dispose();
                    }


                    Console.CursorVisible = true;
                    var medi = TimeSpan.FromMilliseconds(MathA.FiveNumberSummary(rowDatas.Select(a => a.waitTime.TotalMilliseconds))[2]);
                    var gps = rowDatas.OrderBy(o => o.dateTime1).Where(w => w.dateTime1.Day == w.dateTime2.Day && w.dateTime2.Hour >= 7 && w.dateTime2.Hour <= 17).ToList();
                    var gpd = gps.GroupBy(g => g.dateTime1.Year + "/"+g.dateTime1.Month + "/" + g.dateTime1.Day);
                    Console.WriteLine(" ！！ 解析完了 ！！ " + "{0, 4:d0}%   ", 100);
                    Console.WriteLine("");
                    Console.WriteLine("↓*↓*↓*↓*↓*↓*↓*↓*↓*↓*↓*↓*↓*↓*↓*↓*↓*↓*↓*↓*");
                    Console.WriteLine("");
                    Console.WriteLine($"日付別({gpd.First().First().dateTime1.Year}/{gpd.First().First().dateTime1.Month})" +
                        $",待ち時間,0.0104166666666667,0.0625");
                    Console.WriteLine($" ,件数/日,中央値(日),平均値,中央値(月)");
                    List<RowData> rowDatas2 = new List<RowData>();
                    foreach (var item in gpd)
                    {
                        var wtime = item.GroupBy(g => g.pID).Select(s => s.First()).ToList();
                        rowDatas2.AddRange(wtime);
                        var wtime2 = wtime.Select(s => s.waitTime.TotalMilliseconds);
                        var StatMean = TimeSpan.FromMilliseconds(MathA.Mean(wtime2));
                        var five = MathA.FiveNumberSummary(wtime2);
                        var StatMedian = TimeSpan.FromMilliseconds(five[2]);
                        Console.WriteLine($"{item.First().dateTime1.Day}日({Function.DoWeekName(item.First().dateTime1.DayOfWeek)}),{wtime2.Count()},{StatMedian.ToString(@"hh\:mm\:ss")},{StatMean.ToString(@"hh\:mm\:ss")},{medi.ToString(@"hh\:mm\:ss")}");
                    }

                    int defd = 27 - gpd.Count();
                    if(defd > 0)
                    {
                        for (int i = 0; i < defd; i++)
                        {
                            Console.WriteLine("Adjustment Row");
                        }
                    }

                    var gpw = rowDatas2.Where(w=> w.dateTime2.Hour >= 7 && w.dateTime2.Hour <= 17).GroupBy(g=>g.dateTime2.DayOfWeek).OrderBy(o=>o.Key);
                    double tenm = new TimeSpan(0, 10, 0).TotalMilliseconds;

                    Console.WriteLine("");
                    Console.WriteLine($"曜日別({gpw.First().First().dateTime1.Year}/{gpw.First().First().dateTime1.Month}),待ち時間,0.00347222222222222");
                    Console.WriteLine($" ,10分以内(%),10分超え(%),中央値,平均値,件数/日");
                    foreach (var item in gpw)
                    {
                        var wdgc = item.GroupBy(g => g.dateTime1.Day).Count();
                        var wtime = item.Select(s => s.waitTime.TotalMilliseconds);
                        var StatMean = TimeSpan.FromMilliseconds(MathA.Mean(wtime));
                        var five = MathA.FiveNumberSummary(wtime);
                        var StatMedian = TimeSpan.FromMilliseconds(five[2]);
                        double tep = ((double)item.Where(w => w.waitTime.TotalMilliseconds <= tenm).Count() / (double)wtime.Count());
                        Console.WriteLine($"{Function.DoWeekName(item.Key)},{tep},{(1-tep)},{StatMedian.ToString(@"hh\:mm\:ss")},{StatMean.ToString(@"hh\:mm\:ss")},{item.Count()/ wdgc}");
                    }
                    int defw = 7 - gpw.Count();
                    if (defw > 0)
                    {
                        for (int i = 0; i < defw; i++)
                        {
                            Console.WriteLine("Adjustment Row");
                        }
                    }

                    var gph = rowDatas2.Where(w => w.dateTime2.Hour >= 7 && w.dateTime2.Hour <= 17).GroupBy(g => g.dateTime2.Hour).OrderBy(o => o.Key);

                    Console.WriteLine("");
                    Console.WriteLine($"時間別({gph.First().First().dateTime1.Year}/{gph.First().First().dateTime1.Month}),待ち時間,件数");
                    Console.WriteLine($" ,10分以内,10分超え,中央値,平均値,件数/時間,10分以内(%)");
                    
                    foreach (var itemh in gph)
                    {
                        var wtime = itemh.Select(s => s.waitTime.TotalMilliseconds);
                        var StatMean = TimeSpan.FromMilliseconds(MathA.Mean(wtime));
                        var five = MathA.FiveNumberSummary(wtime);
                        var StatMedian = TimeSpan.FromMilliseconds(five[2]);
                        double tep = ((double)itemh.Where(w => w.waitTime.TotalMilliseconds <= tenm).Count() / (double)wtime.Count());
                        Console.WriteLine($"{itemh.Key}:00,{itemh.Count()*tep},{itemh.Count()*(1-tep)},{StatMedian.ToString(@"hh\:mm\:ss")},{StatMean.ToString(@"hh\:mm\:ss")},{itemh.Count()},{tep}");
                    }

                    rowDatas = null;
                    rowDatas2 = null;

                    Console.WriteLine("");
                    Console.WriteLine("↑*↑*↑*↑*↑*↑*↑*↑*↑*↑*↑*↑*↑*↑*↑*↑*↑*↑*↑*↑*");

                    Console.WriteLine("");
                    Console.WriteLine("完了しました、出力をコピーしてExcelに貼り付けてください");
                    Console.WriteLine("終了するには何かキーを押してください");

                    Console.ReadLine();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("!! Error !!");
                    Console.WriteLine(ex.Message);
                    Console.WriteLine("");
                    Console.WriteLine("終了するには何かキーを押してください");
                    Console.ReadLine();
                }

            }
            else
            {
                return;
            }
        }
        
    }
    public class RowData
    {
        public double pID;
        public DateTime dateTime1;
        public DateTime dateTime2;
        public TimeSpan waitTime { get { return dateTime2 - dateTime1; } }
    }
    public static class Function
    {
        public static string DoWeekName(DayOfWeek w)
        {
            string ret = "";
            switch (w)
            {
                case DayOfWeek.Sunday:
                    ret = "日";
                    break;
                case DayOfWeek.Monday:
                    ret = "月";
                    break;
                case DayOfWeek.Tuesday:
                    ret = "火";
                    break;
                case DayOfWeek.Wednesday:
                    ret = "水";
                    break;
                case DayOfWeek.Thursday:
                    ret = "木";
                    break;
                case DayOfWeek.Friday:
                    ret = "金";
                    break;
                case DayOfWeek.Saturday:
                    ret = "土";
                    break;
                default:
                    break;
            }
            return ret;
        }
    }

}
