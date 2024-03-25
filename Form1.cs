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

namespace excel_sorting
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            comboBox1.SelectedItem = comboBox1.Items[0];
        }
        private void buttonOpenXlsx_Click_1(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Xlsx files (*.xlsx)|*.xlsx";
            openFileDialog1.FileName = "";

            DialogResult dr = openFileDialog1.ShowDialog();

            if (dr != DialogResult.No && dr != DialogResult.Cancel && dr != DialogResult.Abort && dr != DialogResult.None)
            {
                try
                {
                    string filename = openFileDialog1.FileName;

                    ReadCSVData(filename);

                }
                catch (System.IO.IOException)
                {
                    MessageBox.Show("Закройте файл!");
                }
            }
        }
        
        public void ReadCSVData(string csvFileName)
        {
            string sheetName = "Sheet 1";
            var csvFilereader = new System.Data.DataTable();
            csvFilereader = ReadExcel(csvFileName, sheetName);

            if (csvFilereader != null)
            {
                GetSortedTable(csvFilereader);
            }
        }

        private System.Data.DataTable ReadExcel(string fileName, string sheetName)
        {
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (var pck = new OfficeOpenXml.ExcelPackage(new FileInfo(fileName)))
            {
                var ws = pck.Workbook.Worksheets[0];

                System.Data.DataTable tbl = new System.Data.DataTable();

                if (ws.Dimension != null)
                {
                    // Создаем столбцы таблицы на основе названий ячеек в первой строке Excel.
                    foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    {
                        tbl.Columns.Add(firstRowCell.Text);
                    }

                    try
                    {
                        // Добавляем данные в таблицу.
                        for (int rowNumber = 2; rowNumber <= ws.Dimension.End.Row; rowNumber++)
                        {
                            var row = ws.Cells[rowNumber, 1, rowNumber, ws.Dimension.End.Column];
                            DataRow newRow = tbl.Rows.Add();
                            foreach (var cell in row)
                            {
                                newRow[cell.Start.Column - 1] = cell.Value;
                            }
                        }

                        return tbl;
                    }
                    catch (System.Data.ConstraintException ex)
                    {
                        MessageBox.Show(ex.Message, "Ошибка");
                        return null;
                    }
                }
                else
                {
                    MessageBox.Show("Файл пустой");
                    return null;
                }
            }
        }

        public void GetSortedTable(DataTable dt)
        {
            DataTable sortedDt = new DataTable();

            try
            {
                if (comboBox1.SelectedItem.ToString() == "по возрастанию")
                {
                    //сортируем таблицу по возрастанию по первому столбцу
                    sortedDt = dt.AsEnumerable().OrderBy(x => x.Field<object>(dt.Columns[0])).CopyToDataTable();
                }
                else if (comboBox1.SelectedItem.ToString() == "по убыванию")
                {
                    //сортируем таблицу по убыванию по первому столбцу
                    sortedDt = dt.AsEnumerable().OrderByDescending(x => x.Field<object>(dt.Columns[0])).CopyToDataTable();
                }

                //сохраняем в новый документ
                saveFileDialog1.Filter = "Xlsx files (*.xlsx)|*.xlsx";

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string filename1 = saveFileDialog1.FileName;

                        using (var package = new ExcelPackage())
                        {
                            var worksheet = package.Workbook.Worksheets.Add("SortedData");
                            worksheet.Cells["A1"].LoadFromDataTable(sortedDt, true);
                            FileInfo fi = new FileInfo(filename1);
                            package.SaveAs(fi);
                            package.Dispose();
                        }
                    }
                    catch (System.InvalidOperationException)
                    {
                        MessageBox.Show("Закройте файл!");
                    }
                }
            }
            catch (System.InvalidOperationException)
            {
                MessageBox.Show("Источник не содержит строк");
            }
        }
    }
}
