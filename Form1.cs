using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelPoC
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        const int DATACNT = 5000;

        private void button1_Click(object sender, EventArgs e)
        {
            var ExcelApp = new Excel.Application();
            ExcelApp.Visible = false;

            var wb = ExcelApp.Workbooks.Add();
            var ws1 = wb.Sheets[1];
            ws1.Select(Type.Missing);
            
            /* 各セルにDATACNTで指定している行までで値を配置する。
             * A列: 序数
             * B列: ランダム数
             */
            for (int i = 1; i < DATACNT; i++)
            {
                var rgn = ws1.Cells[i, 1];
                rgn.Value2 = i;

                var rgn2 = ws1.Cells[i, 2];
                var val = new Random().NextDouble();
                
                rgn2.Value2 = val;
                label_progress.Text = "生成中" + i.ToString();
            }

            label_progress.Text = "データ生成完了";

            /* chart object を生成し、グラフを表示する */
            var chartObjects = ws1.ChartObjects();
            var chartObject = chartObjects.Add(0, 0, 1000, 200);
            var chart = chartObject.Chart;

            /* グラフの設定 */
            chart.ChartType = Excel.XlChartType.xlXYScatterLinesNoMarkers;  // 散布図折れ線
            chart.HasLegend = false;                                        // 凡例を非表示にする
            string ran = "A1:B" + DATACNT.ToString();                       // range 指定用の文字列
            chart.SetSourceData(ws1.Range(ran));                            // chart 描画のデータ元を挿入する
            
            
            label_progress.Text = "Chart生成完了";

            /* 保存と終了 */
            wb.SaveAs("test.xlsx");
            wb.Close(false);
            ExcelApp.Quit();
            label_progress.Text = "保存完了";
        }
    }
}
