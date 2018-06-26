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

        private string ModuleIndex;
        public static class Data_path
        {
            public static string Text { set; get; }
        }
        
        public Form1()
        {
            InitializeComponent();
                        
        }

        public static void Add_New_Sheet_type2( Excel.Workbook excelappworkbook_fun)
        {
            Excel.Sheets excelsheets_fun;
            Excel.Worksheet excelworksheet_fun;

        int sheetscount = excelappworkbook_fun.Sheets.Count;
            excelappworkbook_fun.Worksheets[2].Copy(After: excelappworkbook_fun.Worksheets[sheetscount - 1]);
            excelappworkbook_fun.Worksheets[sheetscount].Name = sheetscount;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[1].ColumnWidth = 0.92;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[2].ColumnWidth = 2;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[3].ColumnWidth = 2;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[4].ColumnWidth = 2.86;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[5].ColumnWidth = 4.43;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[6].ColumnWidth = 0.92;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[7].ColumnWidth = 9.43;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[8].ColumnWidth = 7;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[9].ColumnWidth = 4.43;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[10].ColumnWidth = 32.43;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[11].ColumnWidth = 1.86;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[12].ColumnWidth = 1.86;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[13].ColumnWidth = 1.86;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[14].ColumnWidth = 1.86;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[15].ColumnWidth = 1.86;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[16].ColumnWidth = 1.86;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[17].ColumnWidth = 1.86;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[18].ColumnWidth = 1.86;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[19].ColumnWidth = 1.86;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[20].ColumnWidth = 1.86;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[21].ColumnWidth = 1.86;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[22].ColumnWidth = 1.86;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[23].ColumnWidth = 1.86;
            excelsheets_fun = excelappworkbook_fun.Worksheets;
            //Добавляем номер новой странице
            excelworksheet_fun = (Excel.Worksheet)excelsheets_fun.get_Item(sheetscount);
            excelworksheet_fun.get_Range("S67","U68").UnMerge();
            excelworksheet_fun.Cells[67, 19] = sheetscount;
            excelworksheet_fun.get_Range("S67", "U68").Merge();
            //Добавляем номер последней странице
            excelworksheet_fun = (Excel.Worksheet)excelsheets_fun.get_Item(sheetscount+1);
            excelworksheet_fun.get_Range("R65", "R66").UnMerge();
            excelworksheet_fun.Cells[65, 18] = sheetscount+1;
            excelworksheet_fun.get_Range("R65", "R66").Merge();
            //Добавляем номер 2 странице
            excelworksheet_fun = (Excel.Worksheet)excelsheets_fun.get_Item(2);
            excelworksheet_fun.get_Range("S67", "U68").UnMerge();
            excelworksheet_fun.Cells[67, 19] = 2;
            excelworksheet_fun.get_Range("S67", "U68").Merge();
            //Добавляем номер 1 странице
            excelworksheet_fun = (Excel.Worksheet)excelsheets_fun.get_Item(1);
            excelworksheet_fun.get_Range("O62", "Q62").UnMerge();
            excelworksheet_fun.Cells[62, 15] = 1;
            excelworksheet_fun.get_Range("O62", "Q62").Merge();
            //Добавляем количество листов 1 странице
            excelworksheet_fun = (Excel.Worksheet)excelsheets_fun.get_Item(1);
            excelworksheet_fun.get_Range("R62", "U62").UnMerge();
            excelworksheet_fun.Cells[62, 18] = sheetscount + 1;
            excelworksheet_fun.get_Range("R62", "U62").Merge();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if(excelapp_ref != null)
            {

                excelapp_ref.DefaultSaveFormat = Excel.XlFileFormat.xlExcel12;
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

                    excelappworkbook_ref.SaveAs(Data_path.Text + name_file + " ПЭ3.xls",  //object Filename
                        Excel.XlFileFormat.xlExcel12,          //object FileFormat
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
                    //Закрываем итоговый файл
                    excelappworkbook_ref.Close(false, Type.Missing, Type.Missing);
                    excelapp_ref.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelapp_ref);
                    excelapp_ref = null;
                    excelappworkbook_ref = null;
                    excelappworkbooks_ref = null;
                    //Закрываем файл с первичным BOM
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
            // получаем выбранный BOM файл
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

            // Открываем эталонный файл
            excelapp_ref = new Excel.Application();
            excelapp_ref.Visible = true;
            excelappworkbooks_ref = excelapp_ref.Workbooks;

            excelappworkbook_ref = excelappworkbooks_ref.Open(Application.StartupPath.ToString() + "\\BOM_reference.xlsx",
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing);

            excelsheets_ref = excelappworkbook_ref.Worksheets;
            excelworksheet_ref = (Excel.Worksheet)excelsheets_ref.get_Item(1);
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
            
            excelsheets_ref = excelappworkbook_ref.Worksheets;
            excelworksheet_ref = (Excel.Worksheet)excelsheets_ref.get_Item(1);

            /*Excel.Worksheet newWorksheet;
            newWorksheet = (Excel.Worksheet)excelsheets_ref.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            newWorksheet.Name = "TEst";*/

            // Копирование листов
            //excelappworkbook_ref.Worksheets[1].Copy(excelappworkbook.Worksheets[1]);
            /*int i = 3;
            int 
            Add_New_Sheet(i, excelappworkbook_ref);
            i = 4;
            Add_New_Sheet(i, excelappworkbook_ref);
            
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
            */


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

        private void button4_Click(object sender, EventArgs e)
        {
            Add_New_Sheet_type2(excelappworkbook_ref);

        }
    }
}
