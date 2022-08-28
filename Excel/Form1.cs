using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using excel = Microsoft.Office.Interop.Excel;

namespace Excel
{
    public partial class Form1 : Form
    {
        excel.Application application;
        public Form1()
        {
            InitializeComponent();
            application = new excel.Application();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            excel.Workbook workbook = application.Workbooks.Add();
            try
            {
                excel.Worksheet sheet = application.Worksheets[1];
                sheet.Cells[2, 2] = "Пример";
                sheet.Range["A1"].Value = "Снова пример";
                excel.Range begin = sheet.Cells[4, 1];
                excel.Range end = sheet.Cells[4, 5];
                excel.Range range = sheet.Range[begin, end];
                range.Interior.Color = ColorTranslator.ToOle(Color.LightYellow);
                workbook.SaveAs("C:\\1\\1.xlsx");
                workbook.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                workbook.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            excel.Workbook workbook = application.Workbooks.Add();
            try
            {
                excel.Worksheet sheet = application.Worksheets[1];
                for (int i = 1; i <= 10; i++)
                {
                    sheet.Cells[i, 1] = i;
                }
                excel.Range begin = sheet.Cells[1, 1];
                excel.Range end = sheet.Cells[10, 1];
                excel.Range range = sheet.Range[begin, end];
                for (int i = 1; i <= 10; i++)
                {
                    sheet.Cells[i, 2].Formula = String.Format("=SUM({0})", range.Address);
                    //sheet.Cells[i, 2].FormulaHidden = true;
                    //sheet.Cells[i, 2].Calculate();
                }
                workbook.SaveAs("C:\\1\\1.xlsx");
                workbook.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                workbook.Close();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            excel.Workbook workbook = application.Workbooks.Add();
            try
            {
                excel.Series series;
                excel.Worksheet sheet = application.Worksheets[1];
                Random random = new Random();
                for (int i = 1; i <= 10; i++)
                {
                    sheet.Cells[i, 1] = random.Next(100);
                }
                excel.Range begin = sheet.Cells[1, 1];
                excel.Range end = sheet.Cells[10, 1];
                excel.Range range = sheet.Range[begin, end];
                
                excel.Chart chart = workbook.Charts.Add();
                chart.ChartType = excel.XlChartType.xlLineMarkers;

                //Берём координаты первого графика
                series = chart.SeriesCollection(1);
                series.Values = range;

                chart.Activate();
                chart.Location(excel.XlChartLocation.xlLocationAsObject, "Лист1");
                sheet.Shapes.Item(1).Left = 100;
                sheet.Shapes.Item(1).Top = 100;
                workbook.SaveAs("C:\\1\\1.xlsx");
                workbook.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                workbook.Close();
            }
        }

        //TODO: Сделать крест и график
        private void button4_Click(object sender, EventArgs e)
        {
            excel.Workbook workbook = application.Workbooks.Add();
            int count = Convert.ToInt32(textBox1.Text);
            if (textBox1.Text != null)
            {
                try
                {
                    //Крест
                    excel.Worksheet sheet = application.Worksheets[1];
                    excel.Range begin = sheet.Cells[1, 1];
                    excel.Range end = sheet.Cells[count, count];
                    excel.Range range = sheet.Range[begin, end];
                    Random random = new Random();
                    //Крест
                    range.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                    int c = count;
                    for (int i = 1; i <= count; i++)
                    {
                        sheet.Cells[i, i].Interior.Color = ColorTranslator.ToOle(Color.Red);
                        sheet.Cells[i, c--].Interior.Color = ColorTranslator.ToOle(Color.Red);
                    }

                    //График пирог
                    for (int i = 1; i <= count; i++)
                    {
                        sheet.Cells[i, count + 1] = random.Next(200);
                    }
                    excel.Range begin1 = sheet.Cells[1, count + 1];
                    excel.Range end1 = sheet.Cells[count, count + 1];
                    excel.Range range1 = sheet.Range[begin1, end1];
                    excel.Chart chart = workbook.Charts.Add();
                    chart.ChartType = excel.XlChartType.xlPie;

                    // series = (excel.Series)chart.SeriesCollection(1);
                    //series.Values = range1;
                    chart.SetSourceData(range1, Type.Missing);

                    chart.Activate();
                    chart.Location(excel.XlChartLocation.xlLocationAsObject, "Лист1");
                    sheet.Shapes.Item(1).Left = 0;
                    sheet.Shapes.Item(1).Top = count * 15;
                    sheet.Shapes.Item(1).Width = count * 48;

                    workbook.SaveAs("C:\\1\\1.xlsx");
                    workbook.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    workbook.Close();
                } 
            }
        }
    }
}
