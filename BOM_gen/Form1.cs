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


        Excel.Sheets excelsheets_fin;
        Excel.Worksheet excelworksheet_fin;

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

        public static void Add_New_Sheet_type2( Excel.Workbook excelappworkbook_fun, string indexIzd)
        {
            Excel.Sheets excelsheets_fun;
            Excel.Worksheet excelworksheet_fun;

        int sheetscount = excelappworkbook_fun.Sheets.Count;
            excelappworkbook_fun.Worksheets[2].Copy(After: excelappworkbook_fun.Worksheets[sheetscount - 1]);
            excelappworkbook_fun.Worksheets[sheetscount].Name = sheetscount;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[1].ColumnWidth = 0.92;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[2].ColumnWidth = 2;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[3].ColumnWidth = 2;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[4].ColumnWidth = 3.29;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[5].ColumnWidth = 4.43;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[6].ColumnWidth = 0.92;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[7].ColumnWidth = 9.43;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[8].ColumnWidth = 6;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[9].ColumnWidth = 4.43;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[10].ColumnWidth = 32.43;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[11].ColumnWidth = 1.86;
            excelappworkbook_fun.Worksheets[sheetscount].Columns[12].ColumnWidth = 2.43;
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
            excelworksheet_fun = (Excel.Worksheet)excelsheets_fun.get_Item(sheetscount);
            for (int i = 2; i<= 73; i++)
            {
                for(int j = 4; j<= 17; j++)
                {
                    excelworksheet_fun.Cells[i, j] = null;
                }
            }
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

            //Добавляем индексы на все листы
            excelworksheet_fun = (Excel.Worksheet)excelsheets_fun.get_Item(1);
            excelworksheet_fun.Range[excelworksheet_fun.Cells[58, 10], excelworksheet_fun.Cells[60, 21]].UnMerge();
            excelworksheet_fun.Cells[58, 10] = indexIzd + " ПЭ3";
            excelworksheet_fun.Range[excelworksheet_fun.Cells[58, 10], excelworksheet_fun.Cells[60, 21]].Merge();

            excelworksheet_fun = (Excel.Worksheet)excelsheets_fun.get_Item(2);
            excelworksheet_fun.Range[excelworksheet_fun.Cells[74, 10], excelworksheet_fun.Cells[76, 17]].UnMerge();
            excelworksheet_fun.Cells[74, 10] = indexIzd + " ПЭ3";
            excelworksheet_fun.Range[excelworksheet_fun.Cells[74, 10], excelworksheet_fun.Cells[76, 17]].Merge();

            excelworksheet_fun = (Excel.Worksheet)excelsheets_fun.get_Item(sheetscount);
            excelworksheet_fun.Range[excelworksheet_fun.Cells[74, 10], excelworksheet_fun.Cells[76, 17]].UnMerge();
            excelworksheet_fun.Cells[74, 10] = indexIzd + " ПЭ3";
            excelworksheet_fun.Range[excelworksheet_fun.Cells[74, 10], excelworksheet_fun.Cells[76, 17]].Merge();

            excelworksheet_fun = (Excel.Worksheet)excelsheets_fun.get_Item(sheetscount + 1);
            excelworksheet_fun.Range[excelworksheet_fun.Cells[64, 13], excelworksheet_fun.Cells[66, 17]].UnMerge();
            excelworksheet_fun.Cells[64, 13] = indexIzd + " ПЭ3";
            excelworksheet_fun.Range[excelworksheet_fun.Cells[64, 13], excelworksheet_fun.Cells[66, 17]].Merge();
        }


        //Функция возвращает правильно сформированные позиционные обозначения
        public static string[] DesignatorForm(string str_Elements)
        {
            string LabelGost = null;
            string[] NumberGost = new string[500];
            string[] NumberGostNew = new string[500];
            int[] NumberGostInt = new int[500]; // массив номеров элементов
            int[] NumberSim = new int[500];     // Число символов в строках
            string[] resultGost = new string[500];
            string[] OutputNumber = new string[500];

            
            string[] words = str_Elements.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);

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

                }
                NumberGostInt[r] = Convert.ToInt32(NumberGost[r]);
            }
            
            int startGost = 1;
            int temp = 0;
            for (int r = 1; r <= words.Length - 1; r++)
            {

                if (NumberGostInt[r] - NumberGostInt[r - 1] == 1)
                {
                    startGost++;
                }
                else
                {
                    if (startGost == 1)
                    {
                        startGost = 1;
                        NumberGostNew[temp] = LabelGost + NumberGostInt[r - 1] + ",";
                        NumberSim[temp] = (int)Math.Log10(NumberGostInt[r - 1]) + LabelGost.Length + 2;
                        temp++;
                    }
                    else
                    {
                        if (startGost == 2)
                        {
                            NumberGostNew[temp] = LabelGost + NumberGostInt[r - startGost] + ",";
                            NumberSim[temp] = (int)Math.Log10(NumberGostInt[r - startGost]) + LabelGost.Length + 2;

                            temp++;
                            NumberGostNew[temp] = LabelGost + NumberGostInt[r - 1] + ",";
                            NumberSim[temp] = (int)Math.Log10(NumberGostInt[r - 1]) + LabelGost.Length + 2;
                            startGost = 1;
                            temp++;
                        }
                        else
                        {

                            NumberGostNew[temp] = LabelGost + NumberGostInt[r - startGost ] + "-";
                            NumberSim[temp] = (int)Math.Log10(NumberGostInt[r - startGost ]) + LabelGost.Length + 2;
                            temp++;
                            NumberGostNew[temp] = LabelGost + NumberGostInt[r - 1] + ",";
                            NumberSim[temp] = (int)Math.Log10(NumberGostInt[r - 1]) + LabelGost.Length + 2;
                            startGost = 1;
                            temp++;
                        }
                    }
                }                
            }
            if (startGost == 1)
            {
                NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length - 1];
                NumberSim[temp] = (int)Math.Log10(NumberGostInt[words.Length - 1]) + LabelGost.Length + 1;
            }
            else
            {
                if (startGost == 2)
                {
                    NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length - startGost] + ",";
                    NumberSim[temp] = (int)Math.Log10(NumberGostInt[words.Length - startGost]) + LabelGost.Length + 2;
                    temp++;
                    NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length - 1];
                    NumberSim[temp] = (int)Math.Log10(NumberGostInt[words.Length - 1]) + LabelGost.Length + 1;
                }
                else
                {
                    NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length - startGost] + "-";
                    NumberSim[temp] = (int)Math.Log10(NumberGostInt[words.Length - startGost]) + LabelGost.Length + 2;
                    temp++;
                    NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length - 1];
                    NumberSim[temp] = (int)Math.Log10(NumberGostInt[words.Length - 1]) + LabelGost.Length + 1;
                }
            }
            
            // Формируем окончательный вид
            string tempSumm = "";
            int tempOut = 0;
            for (int hh = 0; hh <= temp; hh++)
            {
                if ((tempSumm.Length + NumberSim[hh]) <= 8)
                {
                    tempSumm = tempSumm + NumberGostNew[hh];
                    if (hh == temp)
                    {
                        OutputNumber[tempOut] = tempSumm;
                    }
                }
                else
                {
                    OutputNumber[tempOut] = tempSumm;
                    tempSumm = NumberGostNew[hh];
                    tempOut++;
                    if (hh == temp)
                    {
                        OutputNumber[tempOut] = tempSumm;
                    }
                }

            }
            return OutputNumber;
        }

        //Функция возвращает кол-во строк позиционных обозначений
        public static int DesignatorFormNumber(string str_Elements)
        {
            string LabelGost = null;
            string[] NumberGost = new string[500];
            string[] NumberGostNew = new string[500];
            int[] NumberGostInt = new int[500]; // массив номеров элементов
            int[] NumberSim = new int[500];     // Число символов в строках
            string[] resultGost = new string[500];
            string[] OutputNumber = new string[500];


            string[] words = str_Elements.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);

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

                }
                NumberGostInt[r] = Convert.ToInt32(NumberGost[r]);
            }

            int startGost = 1;
            int temp = 0;
            for (int r = 1; r <= words.Length - 1; r++)
            {

                if (NumberGostInt[r] - NumberGostInt[r - 1] == 1)
                {
                    startGost++;
                }
                else
                {
                    if (startGost == 1)
                    {
                        startGost = 1;
                        NumberGostNew[temp] = LabelGost + NumberGostInt[r - 1] + ",";
                        NumberSim[temp] = (int)Math.Log10(NumberGostInt[r - 1]) + LabelGost.Length + 2;
                        temp++;
                    }
                    else
                    {
                        if (startGost == 2)
                        {
                            NumberGostNew[temp] = LabelGost + NumberGostInt[r - startGost] + ",";
                            NumberSim[temp] = (int)Math.Log10(NumberGostInt[r - startGost]) + LabelGost.Length + 2;

                            temp++;
                            NumberGostNew[temp] = LabelGost + NumberGostInt[r - 1] + ",";
                            NumberSim[temp] = (int)Math.Log10(NumberGostInt[r - 1]) + LabelGost.Length + 2;
                            startGost = 1;
                            temp++;
                        }
                        else
                        {

                            NumberGostNew[temp] = LabelGost + NumberGostInt[r - startGost ] + "-";
                            NumberSim[temp] = (int)Math.Log10(NumberGostInt[r - startGost ]) + LabelGost.Length + 2;
                            temp++;
                            NumberGostNew[temp] = LabelGost + NumberGostInt[r - 1] + ",";
                            NumberSim[temp] = (int)Math.Log10(NumberGostInt[r - 1]) + LabelGost.Length + 2;
                            startGost = 1;
                            temp++;
                        }
                    }
                }
            }
            if (startGost == 1)
            {
                NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length - 1];
                NumberSim[temp] = (int)Math.Log10(NumberGostInt[words.Length - 1]) + LabelGost.Length + 1;
            }
            else
            {
                if (startGost == 2)
                {
                    NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length - startGost] + ",";
                    NumberSim[temp] = (int)Math.Log10(NumberGostInt[words.Length - startGost]) + LabelGost.Length + 2;
                    temp++;
                    NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length - 1];
                    NumberSim[temp] = (int)Math.Log10(NumberGostInt[words.Length - 1]) + LabelGost.Length + 1;
                }
                else
                {
                    NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length - startGost] + "-";
                    NumberSim[temp] = (int)Math.Log10(NumberGostInt[words.Length - startGost]) + LabelGost.Length + 2;
                    temp++;
                    NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length - 1];
                    NumberSim[temp] = (int)Math.Log10(NumberGostInt[words.Length - 1]) + LabelGost.Length + 1;
                }
            }

            // Формируем окончательный вид
            string tempSumm = "";
            int tempOut = 0;
            for (int hh = 0; hh <= temp; hh++)
            {
                if ((tempSumm.Length + NumberSim[hh]) <= 8)
                {
                    tempSumm = tempSumm + NumberGostNew[hh];
                    if (hh == temp)
                    {
                        OutputNumber[tempOut] = tempSumm;
                    }
                }
                else
                {
                    OutputNumber[tempOut] = tempSumm;
                    tempSumm = NumberGostNew[hh];
                    tempOut++;
                    if (hh == temp)
                    {
                        OutputNumber[tempOut] = tempSumm;
                    }
                }

            }
            return tempOut;
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
            string[] tempsStr = new string[500];
            
            excelsheets = excelappworkbook.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            excelcells = excelworksheet.Cells[i, j];
            sStr = Convert.ToString(excelcells.Value2);
           // richTextBox1.AppendText(sStr+" \n");


            while (sStr != null)
            {
               // i = 1;
                excelsheets = excelappworkbook.Worksheets;
                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
                excelcells = excelworksheet.Cells[i, j];
                sStr = Convert.ToString(excelcells.Value2);

                //richTextBox1.AppendText(sStr + " \n");
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
                            //richTextBox1.AppendText(excelworksheet.Cells[i, j].Value + " \n");
                            i++;
                            excelcells = excelworksheet.Cells[i, j];
                            sStr = Convert.ToString(excelcells.Value2);
                            /*richTextBox1.AppendText("sStr = " + sStr + " \n");
                            //int fsdf = DesignatorFormNumber(sStr);
                            if (sStr != null)
                            {
                                tempsStr = DesignatorForm(sStr);
                                //richTextBox2.AppendText("sStr = " + sStr + "+ \n");
                                for (int r = 0; r <= DesignatorFormNumber(sStr); r++)
                                {
                                    richTextBox1.AppendText(tempsStr[r] + " \n");
                                }
                            }*/

                        }
                        break;
                    case "Quantity":
                        i = 2;
                        excelcells = excelworksheet.Cells[i, j];
                        sStr = Convert.ToString(excelcells.Value2);
                        while (sStr != null)
                        {
                            excelworksheet.Cells[i, 14] = excelworksheet.Cells[i, j];
                            //richTextBox1.AppendText(excelworksheet.Cells[i, j].Value + " \n");
                            i++;
                            excelcells = excelworksheet.Cells[i, j];
                            sStr = Convert.ToString(excelcells.Value2);
                        }
                        break;
                    case "ValueName":
                        i = 2;
                        excelcells = excelworksheet.Cells[i, j];
                        sStr = Convert.ToString(excelcells.Value2);
                        int y = 1;
                        while (y <= max_poz)
                        {
                            excelworksheet.Cells[i, 12] = excelworksheet.Cells[i, j];
                            //richTextBox1.AppendText(excelworksheet.Cells[i, j].Value + " \n");
                            i++;
                            excelcells = excelworksheet.Cells[i, j];
                            sStr = Convert.ToString(excelcells.Value2);
                            y++;
                        }
                        break;
                    case "ValueType":
                        i = 2;
                        excelcells = excelworksheet.Cells[i, j];
                        sStr = Convert.ToString(excelcells.Value2);
                        while (sStr != null)
                        {
                            excelworksheet.Cells[i, 11] = excelworksheet.Cells[i, j];
                            //richTextBox1.AppendText(excelworksheet.Cells[i, j].Value + " \n");
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
                            //richTextBox1.AppendText(excelworksheet.Cells[i, j].Value + " \n");
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
                            //richTextBox1.AppendText(excelworksheet.Cells[i, j].Value + " \n");
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
            int number_string = 0;
            int number_sheet = 1;
            // Записываем значения в таблицу по ГОСТ
            
            excelsheets_fin = excelappworkbook_ref.Worksheets;
            excelworksheet_fin = (Excel.Worksheet)excelsheets_fin.get_Item(1);
            
            // Название прибора 
            excelworksheet_fin.Range[excelworksheet_fin.Cells[61, 10], excelworksheet_fin.Cells[63, 11]].UnMerge();
            excelworksheet_fin.Cells[61, 10] = textBox1.Text;
            excelworksheet_fin.Range[excelworksheet_fin.Cells[61, 10], excelworksheet_fin.Cells[63, 11]].Merge();
            // Разработчик
            excelworksheet_fin.Range[excelworksheet_fin.Cells[61, 6], excelworksheet_fin.Cells[61, 7]].UnMerge();
            excelworksheet_fin.Cells[61, 6] = textBox4.Text;
            excelworksheet_fin.Range[excelworksheet_fin.Cells[61, 6], excelworksheet_fin.Cells[61, 7]].Merge();
            // Проверяющий
            excelworksheet_fin.Range[excelworksheet_fin.Cells[62, 6], excelworksheet_fin.Cells[62, 7]].UnMerge();
            excelworksheet_fin.Cells[62, 6] = textBox5.Text;
            excelworksheet_fin.Range[excelworksheet_fin.Cells[62, 6], excelworksheet_fin.Cells[62, 7]].Merge();
            // Утверждающий 
            excelworksheet_fin.Range[excelworksheet_fin.Cells[65, 6], excelworksheet_fin.Cells[65, 7]].UnMerge();
            excelworksheet_fin.Cells[65, 6] = textBox6.Text;
            excelworksheet_fin.Range[excelworksheet_fin.Cells[65, 6], excelworksheet_fin.Cells[65, 7]].Merge();

            int saveNumberString = 0;
            for (int tempi = 2; tempi <= max_poz; tempi++)
            {
                if (number_string <= 17)
                {
                    number_sheet = 1;
                    saveNumberString = number_string;
                    richTextBox1.AppendText("saveNumberString = " + saveNumberString + " \n");
                }
                else
                {
                    if (number_string <= (17 + (24 * (number_sheet - 1) - 1)))
                    {
                        number_sheet = number_sheet;
                        saveNumberString = number_string - (18 + 24 * (number_sheet - 2));
                        richTextBox1.AppendText("saveNumberString = " + saveNumberString + " \n");
                    }
                    else
                    {
                        number_sheet++;
                        saveNumberString = 0;
                        if(number_sheet >= 3)
                        {
                            Add_New_Sheet_type2(excelappworkbook_ref, textBox3.Text);
                        }
                    }
                }

                 
                excelworksheet_fin = (Excel.Worksheet)excelsheets_fin.get_Item(number_sheet);
                
                //Вписываем наименование
                excelcells = excelworksheet.Cells[tempi, 11];
                string elType = Convert.ToString(excelcells.Value2);
                excelcells = excelworksheet.Cells[tempi, 12];
                string elName = Convert.ToString(excelcells.Value2);
                if (elName != null)
                {
                    excelcells = excelworksheet.Cells[tempi, 12];
                    elName = Convert.ToString(excelcells.Value2);
                }
                else
                {
                    excelcells = excelworksheet.Cells[tempi, 15];
                    elName = Convert.ToString(excelcells.Value2);
                }
                excelworksheet_fin.Range[excelworksheet_fin.Cells[(3 * saveNumberString + 2), 7], excelworksheet_fin.Cells[(3 * saveNumberString + 4), 10]].UnMerge();
                excelworksheet_fin.Cells[(3 * saveNumberString + 2), 7] = elType + " " + elName;
                excelworksheet_fin.Range[excelworksheet_fin.Cells[(3 * saveNumberString + 2), 7], excelworksheet_fin.Cells[(3 * saveNumberString + 4), 10]].Merge();
                
                //Вписываем ТУ
                excelcells = excelworksheet.Cells[tempi, 13];
                string elTech = Convert.ToString(excelcells.Value2);              
                excelworksheet_fin.Range[excelworksheet_fin.Cells[(3 * (saveNumberString + 1) + 2), 7], excelworksheet_fin.Cells[(3 * (saveNumberString + 1) + 4), 10]].UnMerge();
                excelworksheet_fin.Cells[(3 * (saveNumberString + 1) + 2), 7] = elTech;
                excelworksheet_fin.Range[excelworksheet_fin.Cells[(3 * (saveNumberString + 1) + 2), 7], excelworksheet_fin.Cells[(3 * (saveNumberString + 1) + 4), 10]].Merge();
                
                //Вписываем кол-во элементов
                excelcells = excelworksheet.Cells[tempi, 14];
                string elQuantity = Convert.ToString(excelcells.Value2);
                excelworksheet_fin.Range[excelworksheet_fin.Cells[(3 * saveNumberString + 2), 11], excelworksheet_fin.Cells[(3 * saveNumberString + 4), 12]].UnMerge();
                excelworksheet_fin.Cells[(3 * saveNumberString + 2), 11] = elQuantity;
                excelworksheet_fin.Range[excelworksheet_fin.Cells[(3 * saveNumberString + 2), 11], excelworksheet_fin.Cells[(3 * saveNumberString + 4), 12]].Merge();
                
                //Вписываем позиционные обозначения
                excelcells = excelworksheet.Cells[tempi, 10];
                string exstr = Convert.ToString(excelcells.Value2);
                int tempsStrNum = 0;
                if (exstr != null)
                {
                    tempsStr = DesignatorForm(exstr);
                    tempsStrNum = DesignatorFormNumber(exstr);

                }                
                int time = 0;
                while(time <= tempsStrNum)
                {
                    if(number_sheet == 1)
                    {
                        if(number_string <= 17)
                        {
                            excelworksheet_fin = (Excel.Worksheet)excelsheets_fin.get_Item(number_sheet);
                            excelworksheet_fin.Range[excelworksheet_fin.Cells[(3 * number_string + 2), 4], excelworksheet_fin.Cells[(3 * number_string + 4), 6]].UnMerge();
                            excelworksheet_fin.Cells[(3 * number_string + 2), 4] = tempsStr[time];
                            excelworksheet_fin.Range[excelworksheet_fin.Cells[(3 * number_string + 2), 4], excelworksheet_fin.Cells[(3 * number_string + 4), 6]].Merge();                            
                            time++;
                            number_string++;
                        }
                        else
                        {
                            number_sheet++;
                            excelworksheet_fin = (Excel.Worksheet)excelsheets_fin.get_Item(number_sheet);
                            excelworksheet_fin.Range[excelworksheet_fin.Cells[(3 * (number_string-18) + 2), 4], excelworksheet_fin.Cells[(3 * (number_string - 18) + 4), 6]].UnMerge();
                            excelworksheet_fin.Cells[(3 * (number_string - 18) + 2), 4] = tempsStr[time];
                            excelworksheet_fin.Range[excelworksheet_fin.Cells[(3 * (number_string-18) + 2), 4], excelworksheet_fin.Cells[(3 * (number_string - 18) + 4), 6]].Merge();
                            time++;
                            number_string++;
                        }
                    } 
                    else
                    {
                        if(number_string <= 18 + (number_sheet-1)*24 )
                        {
                            excelworksheet_fin = (Excel.Worksheet)excelsheets_fin.get_Item(number_sheet);
                            excelworksheet_fin.Range[excelworksheet_fin.Cells[(3 * (number_string - (18 + 24 * (number_sheet - 2))) + 2) , 4], excelworksheet_fin.Cells[(3 * (number_string - (18 + 24 * (number_sheet - 2))) + 4), 6]].UnMerge();
                            excelworksheet_fin.Cells[(3 * (number_string - (18 + 24 * (number_sheet - 2))) + 2) , 4] = tempsStr[time];
                            excelworksheet_fin.Range[excelworksheet_fin.Cells[(3 * (number_string - (18 + 24 * (number_sheet - 2))) + 2), 4], excelworksheet_fin.Cells[(3 * (number_string - (18 + 24 * (number_sheet - 2))) + 4), 6]].Merge();
                            time++;
                            number_string++;
                        }
                        else
                        {
                            number_sheet++;
                            if (number_sheet >= 3)
                            {
                                Add_New_Sheet_type2(excelappworkbook_ref, textBox3.Text);
                            }
                            excelworksheet_fin = (Excel.Worksheet)excelsheets_fin.get_Item(number_sheet);
                            excelworksheet_fin.Range[excelworksheet_fin.Cells[(3 * (number_string - (18 + 24 * (number_sheet - 2))) + 2), 4], excelworksheet_fin.Cells[(3 * (number_string - (18 + 24 * (number_sheet - 2))) + 4) , 6]].UnMerge();
                            excelworksheet_fin.Cells[(3 * (number_string - (18 + 24 * (number_sheet - 2))) + 2) , 4] = tempsStr[time];
                            excelworksheet_fin.Range[excelworksheet_fin.Cells[(3 * (number_string - (18 + 24 * (number_sheet - 2))) + 2), 4], excelworksheet_fin.Cells[(3 * (number_string - (18 + 24 * (number_sheet - 2))) + 4), 6]].Merge();
                            time++;
                            number_string++;

                        }
                    }
                }

                if (exstr != null)
                {
                    if (DesignatorFormNumber(exstr) <= 0)
                    {
                        number_string++;
                    }
                }
                number_string++;

            }
            richTextBox1.AppendText("Кол. элементов = "+max_poz + " \n");
            

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
            Add_New_Sheet_type2(excelappworkbook_ref, textBox3.Text);

        }

        private void button6_Click(object sender, EventArgs e)
        {
            string Test_str = "C1, C20, C21, C52";
            string str_Elements = " C100, C101";
            string LabelGost = null;
            string[] NumberGost = new string[500];
            string[] NumberGostNew = new string[500];
            int[] NumberGostInt = new int[500]; // массив номеров элементов
            int[] NumberSim = new int[500];     // Число символов в строках
            string[] resultGost = new string[500];
            string[] OutputNumber = new string[500];
            //char[] chartest = new char[500];


            string[] words = str_Elements.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);

            

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

                    //richTextBox1.AppendText(OneWord[j] + " \n");
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
                        NumberSim[temp] = (int)Math.Log10(NumberGostInt[r - 1]) + LabelGost.Length + 2;
                        //richTextBox1.AppendText("Вывод "+ NumberGostNew[temp] + " \n");
                        temp++;
                    }
                    else
                    {
                        if (startGost == 2)
                        {
                            NumberGostNew[temp] = LabelGost + NumberGostInt[r - 2] + ",";
                            NumberSim[temp] = (int)Math.Log10(NumberGostInt[r - 2]) + LabelGost.Length + 2;
                            richTextBox1.AppendText("Вывод " + NumberGostNew[temp] + " \n");
                            temp++;
                            NumberGostNew[temp] = LabelGost + NumberGostInt[r - 1] + ",";
                            NumberSim[temp] = (int)Math.Log10(NumberGostInt[r - 1]) + LabelGost.Length + 2;
                            richTextBox1.AppendText("Вывод " + NumberGostNew[temp] + " \n");
                            startGost = 1;
                            temp++;
                        }
                        else
                        {

                            NumberGostNew[temp] = LabelGost + NumberGostInt[r - startGost ] + "-";
                            NumberSim[temp] = (int)Math.Log10(NumberGostInt[r - startGost ]) + LabelGost.Length + 2;
                            richTextBox1.AppendText("Вывод " + NumberGostNew[temp] + " \n");
                            temp++;
                            NumberGostNew[temp] = LabelGost + NumberGostInt[r - 1] + ",";
                            NumberSim[temp] = (int)Math.Log10(NumberGostInt[r - 1]) + LabelGost.Length + 2;
                            richTextBox1.AppendText("Вывод " + NumberGostNew[temp] + " \n");
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
                NumberSim[temp] = (int)Math.Log10(NumberGostInt[words.Length - 1]) + LabelGost.Length + 1;
                richTextBox1.AppendText("Вывод " + NumberGostNew[temp] + " \n");
            }
            else
            {
                if(startGost == 2)
                {
                    NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length - startGost] + ",";
                    NumberSim[temp] = (int)Math.Log10(NumberGostInt[words.Length - startGost]) + LabelGost.Length + 2;
                    richTextBox1.AppendText("Вывод " + NumberGostNew[temp] + " \n");
                    temp++;
                    NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length - 1] ;
                    NumberSim[temp] = (int)Math.Log10(NumberGostInt[words.Length - 1]) + LabelGost.Length + 1;
                    richTextBox1.AppendText("Вывод " + NumberGostNew[temp] + " \n");
                }
                else
                {
                    NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length - startGost] + "-";
                    NumberSim[temp] = (int)Math.Log10(NumberGostInt[words.Length - startGost]) + LabelGost.Length + 2;
                    richTextBox1.AppendText("Вывод " + NumberGostNew[temp] + " \n");
                    temp++;
                    NumberGostNew[temp] = LabelGost + NumberGostInt[words.Length-1];
                    NumberSim[temp] = (int)Math.Log10(NumberGostInt[words.Length - 1]) + LabelGost.Length + 1;
                    richTextBox1.AppendText("Вывод " + NumberGostNew[temp] + " \n");
                }
            }
            for (int r = 0; r <= temp; r++)
            {
                //string teststr = Convert.ToString(NumberGostNew[r]);
                //int tempst = teststr[0].GetLength(0);
                richTextBox1.AppendText(NumberGostNew[r] + "    ------- "+ NumberSim[r]  +  " \n");
                //richTextBox1.AppendText(LabelGost + LabelGost.Length + NumberGostInt[r]+" \n");
            }

            // Формируем окончательный вид
            string tempSumm = "";
            int tempOut = 0;
            for (int hh = 0; hh <= temp; hh++)
            {
                if((tempSumm.Length + NumberSim[hh]) <= 8)
                {
                    tempSumm = tempSumm + NumberGostNew[hh];
                    if( hh == temp)
                    {
                        OutputNumber[tempOut] = tempSumm;
                    }
                }
                else
                {
                    OutputNumber[tempOut] = tempSumm;
                    tempSumm = NumberGostNew[hh];
                    tempOut++;
                    if (hh == temp)
                    {
                        OutputNumber[tempOut] = tempSumm;
                    }
                }

            }
            for (int r = 0; r <= tempOut; r++)
            {
                richTextBox1.AppendText(OutputNumber[r] + " \n");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {

        }
    }
}
