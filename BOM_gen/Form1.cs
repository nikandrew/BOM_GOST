using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BOM_gen
{
    public partial class Form1 : Form
    {
        private Excel.Application excelapp;
        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;

        private Excel.Application excelapp_ref;
        private Excel.Workbooks excelappworkbooks_ref;
        private Excel.Workbook excelappworkbook_ref;

        private Excel.Range excelcells;
        private Excel.Sheets excelsheets;
        private Excel.Worksheet excelworksheet;

        private Excel.Range excelcells_ref;
        private Excel.Sheets excelsheets_ref;
        private Excel.Worksheet excelworksheet_ref;

        private object _missingObj = System.Reflection.Missing.Value;

        public static class Data_path
        {
            public static string Text { set; get; }
        }
        
        public Form1()
        {
            InitializeComponent();
                        
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(excelapp != null)
            {
                
                excelapp.DefaultSaveFormat = Excel.XlFileFormat.xlAddIn8;
                string name_file = textBox1.Text;
                Data_path.Text = textBox2.Text;
                if (Directory.Exists(Data_path.Text))
                {
                    if (File.Exists(Data_path.Text + name_file + " ПЭ3.xls"))
                    {
                        richTextBox1.AppendText("Файл " + Data_path.Text + name_file + " ПЭ3.xls существует и был перезаписан \n");
                    }
                    else
                    {
                        richTextBox1.AppendText("Файл " + Data_path.Text + name_file + " ПЭ3.xls создан и сохранен \n");
                    }                        
                        
                        excelappworkbook.SaveAs(Data_path.Text + name_file + " ПЭ3.xls",  //object Filename
                        Excel.XlFileFormat.xlAddIn8,          //object FileFormat
                                Type.Missing,                       //object Password 
                                Type.Missing,                       //object WriteResPassword  
                                Type.Missing,                       //object ReadOnlyRecommended
                                Type.Missing,                       //object CreateBackup
                                Excel.XlSaveAsAccessMode.xlNoChange,//XlSaveAsAccessMode AccessMode
                                 Excel.XlSaveConflictResolution.xlLocalSessionChanges,                       //object ConflictResolution
                                Type.Missing,                       //object AddToMru 
                                Type.Missing,                       //object TextCodepage
                                Type.Missing,                       //object TextVisualLayout
                                Type.Missing);
                        excelappworkbook.Close(false, Type.Missing, Type.Missing);
                        excelapp.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelapp);
                        excelapp = null;
                        excelappworkbook = null;
                        excelappworkbooks = null;
                        System.GC.Collect();

                        
                }
                else
                {
                    richTextBox1.AppendText( "Каталог " + Data_path.Text + " не найден \n");
                }
            }
            else
            {
                richTextBox1.AppendText("Нет открытых файлов \n");                
            }  
        }
    
       
        private void button5_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            string filename = openFileDialog1.FileName;
            Data_path.Text = Path.GetDirectoryName(filename)+"\\";
            textBox2.Text = Data_path.Text;
            excelapp = new Excel.Application();
            excelapp.Visible = true;
            excelappworkbooks = excelapp.Workbooks;

            excelappworkbook = excelappworkbooks.Open(filename,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing);
            //excelappworkbook = excelappworkbooks[1];
        }

        private void button1_Click(object sender, EventArgs e)
        {
            excelsheets = excelappworkbook.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            excelcells = excelworksheet.get_Range("A1", Type.Missing);
            string sStr = Convert.ToString(excelcells.Value2);
            richTextBox1.AppendText(sStr+" \n");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Открытие файла
            excelapp_ref = new Excel.Application();
            excelapp_ref.Visible = true;
            excelappworkbooks_ref = excelapp_ref.Workbooks;

            excelappworkbook_ref = excelappworkbooks_ref.Open(Application.StartupPath.ToString() + "\\BOM_reference.xlsx",
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing);
            //excelappworkbook_ref = excelappworkbooks_ref[1];
            excelsheets_ref = excelappworkbook_ref.Worksheets;
            excelworksheet_ref = (Excel.Worksheet)excelsheets_ref.get_Item(1);

            /*Excel.Worksheet newWorksheet;
            newWorksheet = (Excel.Worksheet)excelsheets_ref.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            newWorksheet.Name = "TEst";*/

            // Копирование листов
            //excelappworkbook_ref.Worksheets[1].Copy(excelappworkbook.Worksheets[1]);
            int i = 3;

            excelappworkbook_ref.Worksheets[2].Copy(After: excelappworkbook_ref.Worksheets[2]);
            excelappworkbook_ref.Worksheets[i].Name = i;
            excelappworkbook_ref.Worksheets[i].Columns[1].ColumnWidth = 0.92;
            excelappworkbook_ref.Worksheets[i].Columns[2].ColumnWidth = 2;
            excelappworkbook_ref.Worksheets[i].Columns[3].ColumnWidth = 2;
            excelappworkbook_ref.Worksheets[i].Columns[4].ColumnWidth = 2.86;
            excelappworkbook_ref.Worksheets[i].Columns[5].ColumnWidth = 4.43;
            excelappworkbook_ref.Worksheets[i].Columns[6].ColumnWidth = 0.92;
            excelappworkbook_ref.Worksheets[i].Columns[7].ColumnWidth = 9.43;
            excelappworkbook_ref.Worksheets[i].Columns[8].ColumnWidth = 7;
            excelappworkbook_ref.Worksheets[i].Columns[9].ColumnWidth = 4.43;
            excelappworkbook_ref.Worksheets[i].Columns[10].ColumnWidth = 32.43;
            excelappworkbook_ref.Worksheets[i].Columns[11].ColumnWidth = 1.86;
            excelappworkbook_ref.Worksheets[i].Columns[12].ColumnWidth = 1.86;
            excelappworkbook_ref.Worksheets[i].Columns[13].ColumnWidth = 1.86;
            excelappworkbook_ref.Worksheets[i].Columns[14].ColumnWidth = 1.86;
            excelappworkbook_ref.Worksheets[i].Columns[15].ColumnWidth = 1.86;
            excelappworkbook_ref.Worksheets[i].Columns[16].ColumnWidth = 1.86;
            excelappworkbook_ref.Worksheets[i].Columns[17].ColumnWidth = 1.86;
            excelappworkbook_ref.Worksheets[i].Columns[18].ColumnWidth = 1.86;
            excelappworkbook_ref.Worksheets[i].Columns[19].ColumnWidth = 1.86;
            excelappworkbook_ref.Worksheets[i].Columns[20].ColumnWidth = 1.86;
            excelappworkbook_ref.Worksheets[i].Columns[21].ColumnWidth = 1.86;
            excelappworkbook_ref.Worksheets[i].Columns[22].ColumnWidth = 1.86;
            excelappworkbook_ref.Worksheets[i].Columns[23].ColumnWidth = 1.86;



            //xlApp.Visible = true;


            //Закрытие файла
            /*excelappworkbook_ref.Close(false, Type.Missing, Type.Missing);
            excelapp_ref.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelapp_ref);
            excelapp_ref = null;
            excelappworkbook_ref = null;
            excelappworkbooks_ref = null;
            System.GC.Collect();*/
        }
    }
}
