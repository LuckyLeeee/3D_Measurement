using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _3D_Measurement
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void import_Click(object sender, EventArgs e)
        {
            string fileName;
            using(OpenFileDialog openFileDialog = new OpenFileDialog() { Filter= "Objects File |*.obj" })
            {
                if(openFileDialog.ShowDialog()==DialogResult.OK)
                {
                    fileName = Path.GetFileName(openFileDialog.FileName);
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    var package = new ExcelPackage(new FileInfo("FemaleBodyDataSet.xlsx"));
                    ExcelWorksheet workSheet = package.Workbook.Worksheets[0];
                    for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
                    {
                        if (fileName.StartsWith(workSheet.Cells[i, 1].Value.ToString()))
                        {
                            cd.Text = workSheet.Cells[i, 2].Value.ToString();
                            c7.Text = workSheet.Cells[i, 3].Value.ToString();
                            cv.Text = workSheet.Cells[i, 4].Value.ToString();
                            ce.Text = workSheet.Cells[i, 5].Value.ToString();
                            cb.Text = workSheet.Cells[i, 6].Value.ToString();
                            cg.Text = workSheet.Cells[i, 7].Value.ToString();
                            cm.Text = workSheet.Cells[i, 8].Value.ToString();
                        }
                    }
                }  
            }    
            try
            {
                var package = new ExcelPackage(new FileInfo("FemaleBodyDataSet.xlsx"));

                ExcelWorksheet workSheet = package.Workbook.Worksheets[1];

                for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
                {
                    try
                    {
                      //cd.Text = workSheet.Cells[1, 1].Value.ToString();
                    }
                    catch
                    {

                    }
                } 
            }
            catch
            {

            }
        }

    }
}
