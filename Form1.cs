using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Kwerendy
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;

                //zczytanie excela i wype³nienie listy wszytskich wpisów
                using (var package = new ExcelPackage(new System.IO.FileInfo(filePath)))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    for (int row = 2; row < worksheet.Rows.Count(); row++)
                    {
                        MyTools.allEntries.Add(new Entry(
                                worksheet.Cells[row, 1].Text, worksheet.Cells[row, 2].Text, worksheet.Cells[row, 3].Text,
                                worksheet.Cells[row, 4].Text, worksheet.Cells[row, 5].Text, worksheet.Cells[row, 6].Text,
                                worksheet.Cells[row, 7].Text, worksheet.Cells[row, 8].Text, worksheet.Cells[row, 9].Text
                                ));
                    }

                }
                grid1.DataSource = MyTools.allEntries;

            }

        }

        private void btnRegistrations_Click(object sender, EventArgs e)
        {
            var reg = MyTools.allEntries.Where(n => n.Action == "! ! ! rejestracja ! ! !");
            foreach (var entry in reg) MyTools.registrations.Add(entry);
            grid1.DataSource = MyTools.registrations;

        }

        private void butGenerate_Click(object sender, EventArgs e)
        {
            if(MyTools.signatureFilePath == "")
            {
                MessageBox.Show("Aby wygenerowaæ dokumenty, najpierw wczytaj podpis", "Brak podpisu!");
            }
            else
            {
                FolderBrowserDialog saveFolderDialog = new FolderBrowserDialog();
                saveFolderDialog.Description = "Wybierz folder";
                saveFolderDialog.RootFolder = Environment.SpecialFolder.Desktop;

                if (saveFolderDialog.ShowDialog() == DialogResult.OK)
                    GenerateDoc.Generate(saveFolderDialog.SelectedPath, MyTools.registrations);
            }
            
        }

        private void btnSignature_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Pliki graficzne (*.png;*.jpg)|*.png;*.jpg";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
                MyTools.signatureFilePath = openFileDialog.FileName;
        }
    }

}

