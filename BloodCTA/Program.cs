using ClosedXML.Excel;
using MathNet.Numerics.Statistics;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            Console.WriteLine("※抽出要件");
            Console.WriteLine("エクセルファイルの最終シートを読み込みます");
            Console.WriteLine("1～3列に関して、PID,採血受付日時(時刻1),患者認証(時刻2)である事");

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                Console.WriteLine("Select File：" + ofd.FileName);
                Console.WriteLine("ファイルを読み込んでいます、お待ちください。。。");
                Console.WriteLine("※抽出条件");
                Console.WriteLine("採血受付日時(時刻1)と患者認証(時刻2)が同日である");
                Console.WriteLine("患者認証(時刻2)が7時～17時の間");
                Console.WriteLine("が7時～17時の間");
                try
                {
                    XLWorkbook workbook = new XLWorkbook(ofd.FileName);
                    IXLWorksheet worksheet = workbook.Worksheets.Last();
                    int lastRow = worksheet.LastRowUsed().RowNumber();
                    Console.CursorVisible = false;
                    char[] bars = { '／', '―', '＼', '｜' };
                    List<RowData> rowDatas = new List<RowData>();
                    for (int i = 2; i <= lastRow; i++)
                    {
                        Console.Write(bars[i % 4]);
                        Console.Write("{0, 4:d0}%", 100 * (i + 1) / lastRow);
                        Console.SetCursorPosition(0, Console.CursorTop);

                        IXLCell cell = worksheet.Cell(i, 1);
                        double pid = (double)cell.Value;
                        DateTime dt1;
                        DateTime dt2;
                        if (!DateTime.TryParse(worksheet.Cell(i, 2).Value.ToString(), out dt1)) continue;
                        if (!DateTime.TryParse(worksheet.Cell(i, 3).Value.ToString(), out dt2)) continue;
                        rowDatas.Add(new RowData() { pID = pid, dateTime1 = dt1, dateTime2 = dt2 });

                    }
                    Console.CursorVisible = true;
                    var gps = rowDatas.OrderBy(o => o.dateTime1).Where(w => w.dateTime1.Day == w.dateTime2.Day && w.dateTime2.Hour >= 7 && w.dateTime2.Hour <= 17).ToList();
                    var gpd = gps.GroupBy(g => g.dateTime1.Year + "/"+g.dateTime1.Month + "/" + g.dateTime1.Day);

                    Console.WriteLine($"日付,曜日,総数,平均値,中央値");
                    List<RowData> rowDatas2 = new List<RowData>();
                    foreach (var item in gpd)
                    {
                        var wtime = item.GroupBy(g => g.pID).Select(s => s.First()).ToList();
                        rowDatas2.AddRange(wtime);
                        var wtime2 = wtime.Select(s => s.waitTime.TotalMilliseconds);
                        var StatMean = TimeSpan.FromMilliseconds(Statistics.Mean(wtime2));
                        var five = Statistics.FiveNumberSummary(wtime2);
                        var StatMedian = TimeSpan.FromMilliseconds(five[2]);
                        Console.WriteLine($"{item.Key},{Function.DoWeekName(item.First().dateTime1.DayOfWeek)},{wtime2.Count()},{StatMean.ToString(@"hh\:mm\:ss")},{StatMedian.ToString(@"hh\:mm\:ss")}");
                    }
                    var gpw = rowDatas2.GroupBy(g=>g.dateTime1.DayOfWeek).OrderBy(o=>o.Key);
                   
                    Console.WriteLine("");
                    Console.WriteLine($"曜日,総数,平均値,中央値");
                    foreach (var item in gpw)
                    {
                        var wtime = item.Select(s => s.waitTime.TotalMilliseconds);
                        var StatMean = TimeSpan.FromMilliseconds(Statistics.Mean(wtime));
                        var five = Statistics.FiveNumberSummary(wtime);
                        var StatMedian = TimeSpan.FromMilliseconds(five[2]);
                        Console.WriteLine($"{Function.DoWeekName(item.Key)},{item.Count()},{StatMean.ToString(@"hh\:mm\:ss")},{StatMedian.ToString(@"hh\:mm\:ss")}");
                    }
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
