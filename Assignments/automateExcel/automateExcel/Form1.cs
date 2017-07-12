/*
 * Ryan Luig
 * Software Dev Tools
 * July 12, 2017
 * 
 * *************************************************
 * This program is an example showing how to use c# 
 * to automate Excel. A name and 50 random numbers
 * are put into columns and then a chart is created
 * using the random numbers.
 * 
 */ 



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
using System.Reflection;



namespace automateExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range chartRange;
            
           
            int[] rand = new int[50]; 
            Random randNum = new Random();

            object misValue = System.Reflection.Missing.Value;
           

            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)oSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(100, 0, 300, 250);
                Excel.Chart chartPage = myChart.Chart;



                oSheet.Cells[1, 1] = "Ryan Luig";

                for(int i = 0; i < rand.Length; i++)
                {
                    rand[i] = randNum.Next(0,100);
                }

                int offset = 2;
                for (int i = 0; i < 50; i++)
                {
                    oSheet.Cells[i + offset, 2] = rand[i];
         
                }


                chartRange = oSheet.get_Range("B2", "B51");
                chartPage.SetSourceData(chartRange, misValue);
                chartPage.ChartType = Excel.XlChartType.xlLine;
                chartPage.HasTitle = true;
                chartPage.ChartTitle.Text = "Random Numbers";
                chartPage.SeriesCollection(1).Name = "random";

            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }
    }
}
