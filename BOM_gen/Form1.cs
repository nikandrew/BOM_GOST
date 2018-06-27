﻿using System;
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

        // Заголовки граф
        private string sDesignator = "Designator";
        private string sQuantity = "Quantity";
        private string sValueName = "ValueName";
        private string sValueType = "ValueType";
        private string sDeesignItemId = "DeesignItemId";
        private string sTU = "DeesignItemId";

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
            excelworksheet_fun.get_Range("R75","T76").UnMerge();
            excelworksheet_fun.Cells[67, 19] = sheetscount;
            excelworksheet_fun.get_Range("R75", "T76").Merge();
            //Добавляем номер последней странице
            excelworksheet_fun = (Excel.Worksheet)excelsheets_fun.get_Item(sheetscount+1);
            excelworksheet_fun.get_Range("R65", "R66").UnMerge();
            excelworksheet_fun.Cells[65, 18] = sheetscount+1;
            excelworksheet_fun.get_Range("R65", "R66").Merge();
            //Добавляем номер 2 странице
            excelworksheet_fun = (Excel.Worksheet)excelsheets_fun.get_Item(2);
            excelworksheet_fun.get_Range("R75", "T76").UnMerge();
            excelworksheet_fun.Cells[67, 19] = 2;
            excelworksheet_fun.get_Range("R75", "T76").Merge();
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
            int j = 1;
            int i = 1;
            int max_poz = 0;
            string sStr;
            
            excelsheets = excelappworkbook.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            excelcells = excelworksheet.Cells[i, j];
            sStr = Convert.ToString(excelcells.Value2);
            richTextBox1.AppendText(sStr+" \n");

            while (sStr != null)
            {
               // i = 1;
                excelsheets = excelappworkbook.Worksheets;
                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
                excelcells = excelworksheet.Cells[i, j];
                sStr = Convert.ToString(excelcells.Value2);
                richTextBox1.AppendText(sStr + " \n");
                switch (sStr)
                {
                    case "Designator":
                        i = 2;
                        excelcells = excelworksheet.Cells[i, j];
                        sStr = Convert.ToString(excelcells.Value2);
                        while (sStr != null )
                        {
                            max_poz++;
                            excelworksheet.Cells[i, 10] = excelworksheet.Cells[i, j];                            
                            richTextBox1.AppendText(excelworksheet.Cells[i, j].Value + " \n");
                            i++;
                            excelcells = excelworksheet.Cells[i, j];
                            sStr = Convert.ToString(excelcells.Value2);
                        }
                        break;
                    case "Quantity":
                        i = 2;
                        excelcells = excelworksheet.Cells[i, j];
                        sStr = Convert.ToString(excelcells.Value2);
                        while (sStr != null)
                        {
                            excelworksheet.Cells[i, 14] = excelworksheet.Cells[i, j];
                            richTextBox1.AppendText(excelworksheet.Cells[i, j].Value + " \n");
                            i++;
                            excelcells = excelworksheet.Cells[i, j];
                            sStr = Convert.ToString(excelcells.Value2);
                        }
                        break;
                    case "ValueName":
                        i = 2;
                        excelcells = excelworksheet.Cells[i, j];
                        sStr = Convert.ToString(excelcells.Value2);
                        while (sStr != null)
                        {
                            excelworksheet.Cells[i, 12] = excelworksheet.Cells[i, j];
                            richTextBox1.AppendText(excelworksheet.Cells[i, j].Value + " \n");
                            i++;
                            excelcells = excelworksheet.Cells[i, j];
                            sStr = Convert.ToString(excelcells.Value2);
                        }
                        break;
                    case "ValueType":
                        i = 2;
                        excelcells = excelworksheet.Cells[i, j];
                        sStr = Convert.ToString(excelcells.Value2);
                        while (sStr != null)
                        {
                            excelworksheet.Cells[i, 11] = excelworksheet.Cells[i, j];
                            richTextBox1.AppendText(excelworksheet.Cells[i, j].Value + " \n");
                            i++;
                            excelcells = excelworksheet.Cells[i, j];
                            sStr = Convert.ToString(excelcells.Value2);
                        }
                        break;
                    case "DesignItemId":
                        i = 2;
                        excelcells = excelworksheet.Cells[i, j];
                        sStr = Convert.ToString(excelcells.Value2);
                        while (sStr != null)
                        {
                            excelworksheet.Cells[i, 15] = excelworksheet.Cells[i, j];
                            richTextBox1.AppendText(excelworksheet.Cells[i, j].Value + " \n");
                            i++;
                            excelcells = excelworksheet.Cells[i, j];
                            sStr = Convert.ToString(excelcells.Value2);
                        }
                        break;
                    case "ValueTechReq":
                        i = 2;
                        excelcells = excelworksheet.Cells[i, j];
                        sStr = Convert.ToString(excelcells.Value2);
                        while (sStr != null)
                        {
                            excelworksheet.Cells[i, 13] = excelworksheet.Cells[i, j];
                            richTextBox1.AppendText(excelworksheet.Cells[i, j].Value + " \n");
                            i++;
                            excelcells = excelworksheet.Cells[i, j];
                            sStr = Convert.ToString(excelcells.Value2);
                        }
                        break;
                    default:

                        break;
                }
                j++;
                i = 1;
                excelcells = excelworksheet.Cells[i, j];
                sStr = Convert.ToString(excelcells.Value2);
            }
            richTextBox1.AppendText("Кол. элементов = "+max_poz + " \n");
            // Формируем поля
            for (int t = 1; t <= max_poz; t++)
            {
                excelworksheet.Cells[t, 24] = excelworksheet.Cells[t, 10];
                excelworksheet.Cells[t, 22] = excelworksheet.Cells[t, 13];
                excelworksheet.Cells[t, 23] = excelworksheet.Cells[t, 14];
                string elName = Convert.ToString(excelworksheet.Cells[t, 11].Value2);
                string elType = Convert.ToString(excelworksheet.Cells[t, 12].Value2);
                excelworksheet.Cells[t, 21] = elName + " " + elType;
            }

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

        private void button6_Click(object sender, EventArgs e)
        {
            string Test_str = "C1, C20, C21, C52";
            string Test_str_DA = "R1, R20, R21, R52, R53, R54, R55, R90, R91, R92, R93, R94, R95, R96, R97, R98, R99, R100, R101, R102, R103, R104, R105, R107, R108, R109, R110, R111, R112, R113";
            string LabelGost = null;
            string[] NumberGost = new string[500];
            string[] NumberGostNew = new string[500];
            int[] NumberGostInt = new int[500]; // массив номеров элементов
            string resultGost = null;


            string[] words = Test_str_DA.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i <= words.Length - 1; i++)
            {
                richTextBox1.AppendText(words[i] + " \n");
            }

            for (int r = 0; r <= words.Length - 1; r++)
            {
                LabelGost = null;

                char[] OneWord = words[r].ToCharArray();

                for (int j = 0; j <= OneWord.Length - 1; j++)
                {
                    if (char.IsLetter(OneWord[j]))
                    {
                        LabelGost = LabelGost + OneWord[j];
                    }
                    else
                    {
                        if (char.IsDigit(OneWord[j]))
                        {
                            NumberGost[r] = NumberGost[r] + OneWord[j];
                        }
                    }

                    richTextBox1.AppendText(OneWord[j] + " \n");
                }
                NumberGostInt[r] = Convert.ToInt32(NumberGost[r]);
                richTextBox1.AppendText(LabelGost + " \n");
                richTextBox1.AppendText("Элемент [" + r + "] = " + NumberGostInt[r] + " \n");
            }
            //NumberGostNew[0] = LabelGost + NumberGostInt[0];

            int startGost = 1;
            int temp = 0;
            for (int r = 1; r <= words.Length - 1; r++)
            {
                
                if (NumberGostInt[r] - NumberGostInt[r - 1] == 1)
                {
                    startGost++;
                    //richTextBox1.AppendText("startGost++ \n");
                }
                else
                {
                    if (startGost == 1)
                    {
                        startGost = 1;
                        NumberGostNew[temp] = LabelGost + NumberGostInt[r - 1] + ",";
                        //richTextBox1.AppendText("Вывод "+ NumberGostNew[temp] + " \n");
                        temp++;
                    }
                    else
                    {
                        if (startGost == 2)
                        {
                            NumberGostNew[temp] = LabelGost + NumberGostInt[r - 2] +","+ LabelGost + NumberGostInt[r - 1] + ",";
                            //richTextBox1.AppendText("Вывод " + NumberGostNew[temp] + " \n");
                            startGost = 1;
                            temp++;
                        }
                        else
                        {
                            
                            NumberGostNew[temp] = LabelGost + NumberGostInt[r - startGost + 1] + "-" + LabelGost + NumberGostInt[r-1] + ",";
                            //richTextBox1.AppendText("Вывод " + NumberGostNew[temp] + " \n");
                            startGost = 1;
                            temp++;
                        }
                    }
                }
                //richTextBox1.AppendText("r=" + r + ", DA=" + NumberGostInt[r] + " \n");
                //richTextBox1.AppendText(" startGost = " + startGost + " \n");
            }
            if(startGost == 1)
            {
                NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length - 1] ;
            }
            else
            {
                if(startGost == 2)
                {
                    NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length - 1] + "," + LabelGost + NumberGostInt[words.Length] ;
                }
                else
                {
                    NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length - startGost] + "-" + LabelGost + NumberGostInt[words.Length-1] ;
                }
            }
            for (int r = 0; r <= NumberGostNew.Length - 1; r++)
            {
                richTextBox1.AppendText(NumberGostNew[r] + " \n");
            }

            // Формируем окончательный вид
            {

            }
        }
    }
}
