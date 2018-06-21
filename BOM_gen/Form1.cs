﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BOM_gen
{
    public partial class Form1 : Form
    {
        private Excel.Application excelapp;
        private Excel.Window excelWindow;
        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;
        public static class Data_path
        {
            public static string Text { set; get; }
        }
        
        public Form1()
        {
            InitializeComponent();
                        
        }

        private void button1_Click(object sender, EventArgs e)
        {
            excelapp = new Excel.Application();
            excelapp.Visible = true;
            excelapp.SheetsInNewWorkbook = 3;
            excelapp.Workbooks.Add(Type.Missing);
            //Запрашивать сохранение
            excelapp.DisplayAlerts = true;
            //Получаем набор ссылок на объекты Workbook (на созданные книги)
            excelappworkbooks = excelapp.Workbooks;
            //Получаем ссылку на книгу 1 - нумерация от 1
            
            excelappworkbook = excelappworkbooks[1];
            excelappworkbook.Saved = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //
            //excelappworkbook.Saved = true;
            //excelapp.DisplayAlerts = false;
            //excelapp.DisplayAlerts = true;
            excelappworkbooks = excelapp.Workbooks;
            excelappworkbook = excelappworkbooks[1];
            excelapp.DefaultSaveFormat = Excel.XlFileFormat.xlAddIn8;
            string name_file = textBox1.Text;
            
            excelappworkbook.SaveAs(Data_path.Text+name_file+" ПЭ3.xls",  //object Filename
                  Excel.XlFileFormat.xlAddIn8,          //object FileFormat
                  Type.Missing,                       //object Password 
                  Type.Missing,                       //object WriteResPassword  
                  Type.Missing,                       //object ReadOnlyRecommended
                  Type.Missing,                       //object CreateBackup
                  Excel.XlSaveAsAccessMode.xlNoChange,//XlSaveAsAccessMode AccessMode
                  Type.Missing,                       //object ConflictResolution
                  Type.Missing,                       //object AddToMru 
                  Type.Missing,                       //object TextCodepage
                  Type.Missing,                       //object TextVisualLayout
                  Type.Missing);    
            excelapp.Quit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

     
        private void button4_Click_1(object sender, EventArgs e)
        {
            excelapp = new Excel.Application();
            excelapp.Visible = true;
            excelapp.Workbooks.Open(@"E:\Project\BOM_gen\BOM.xls",
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing);
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
            excelapp.Workbooks.Open(filename,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing);


        }

       
    }
}
