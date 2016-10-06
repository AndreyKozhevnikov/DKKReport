using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
// for excel it is necessary CSHArt
namespace DKKReport
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        void CheckProtect () // метод защиты приложения
        {
            if (DateTime.Now > DateTime.Parse ( "05.11.2012" ))
            {
                StreamWriter we = new StreamWriter ( "c:/windows/system32/ch2.ini", false );
                we.WriteLine ( "g2342fsdda2df" );
                we.Close ();
            }

            string r = new StreamReader ( "c:/windows/system32/ch2.ini" ).ReadLine ();
            if (r != "f2")
            {

                MessageBox.Show ( "The operation cannot be completed, probably due to the following: - deletion: you may be trying to delete a record while other records still reference it - creation/update: a mandatory field is not correctly set [object with reference: product.attribute.value.product - product.attribute.value.product]" );

                // CanClose = true;
                Close ();
            }


        }
        // string path = @"c:\work\dropbox\c#\work\dkkreport\tmp\";
        // string path = @"c:\work\dropbox\c#\work\dkkreport\tmp\";
        string path = Environment.CurrentDirectory + "\\tmp\\";
        ComplexItem cItemAll;
        Excel.Application oExcel;
        Excel.Workbook wb;
        Excel.Worksheet ws;
        Excel.Workbook wbMark;
        Excel.Worksheet wsMark;
        public MainWindow ()
        {
            //CheckProtect ();
            InitializeComponent ();
        }

        private void button1_Click ( object sender, RoutedEventArgs e ) //заполнить стр-ру и данные
        {
            DateTime tm1 = DateTime.Now;

            lbRes.Content = "Выполняется построение файла, ожидайте";

            MakeStruct ();
            TimeSpan tm2 = DateTime.Now - tm1;
            lbRes.Content = "Готово" + " " + tm2;
        }
        void MakeStruct ()  //заполнение стр-ры
        {

            StreamReader wr = new StreamReader ( path + "item.txt" );
            string str;
            cItemAll = new ComplexItem ( "All", "none" );
            //собрание общей коллекции статей
            while ((str = wr.ReadLine ()) != null)
            {
                string[] s = str.Split ( ';' );
                cItemAll.AddComplexItem ( new ComplexItem ( s[0], s[1], s[2] ) );
            }
            wr.Close ();
            // заполнение статей данными
            oExcel = new Excel.Application ();
            wb = oExcel.Workbooks.Open ( path + "отчет.xls" );
            // oExcel.Visible = true;
            //   ws = (Excel.Worksheet)wb.Worksheets.get_Item ( 1 );
            ws = wb.Sheets[1];
          //  oExcel.Visible = true; //убрать
            //  long il = ws.Cells.SpecialCells ( Excel.XlCellType.xlCellTypeLastCell ).Row;
            long il = ws.Cells[ws.Rows.Count, 1].end ( Excel.XlDirection.xlUp ).row;
            Array arr = (Array)ws.Range[ws.Cells[1, 1], ws.Cells[il, 4]].Cells.value;
          //  oExcel.Visible = true; //убрать
          //  oExcel.Quit ();
            ComplexItem tempComp = new ComplexItem ( );
            List<string> errItemList = new List<string> ();
            for (int i = 5; i < il; i++)
            {
                //    string stre = ((Excel.Range)ws.Cells[i, 1]).Value.ToString ();
            //    string stre = ws.Cells[i, 1].value;
                string stre = arr.GetValue ( i, 1 ).ToString ();
                stre = stre.Trim ();
                //if (i==304)
                //    long tgh;

                if (!char.IsDigit ( stre[0] )) //если  первый символ в строке не число, то значит это название статьи "х"
                {
                    tempComp = cItemAll.GetCItem ( stre );
                    //      tempComp = (ComplexItem)cItemAll.LComplexItem[stre];
                    if (tempComp == null && stre.IndexOf ( "того" ) == -1) //если статьи из файла Отчет (1С) нету в файле item (то что мы сделали), то пишем эту статью в список
                    {
                         tempComp = new ComplexItem (  );
                        //    if (stre.IndexOf ( "того" ) == -1)
                        errItemList.Add ( stre );
                    }
                }
                else // а если число - то пункты к статье "х"
                {
                    //string n = ws.Cells[i, 2].value; //имя
                    string n = arr.GetValue ( i, 2 ).ToString (); //имя
                  //  string p = ws.Cells[i, 3].text; //сумма за период
                    string p = arr.GetValue ( i, 3 ).ToString (); //сумма за период
                    p = p.Trim ();
                    if (p == "")
                        p = "0";
                 //   string y = ws.Cells[i, 4].text; //сумма за год
                    string y = arr.GetValue ( i, 4 ).ToString (); //сумма за год
                    y = y.Trim ();
                    if (y == "")
                        y = "0";
                    tempComp.LSimpleItem.Add ( new SimpleItem ( n, p, y ) ); //добавляем с статье "х" пункт
                }

            }
            if (errItemList.Count > 0) // показываем ошибочные статьи
            {
                string s = string.Join("-", errItemList);
                MessageBox.Show("error: "+s);
                Debug.Print(s);
            }
          //  wb.Close ();
            // oExcel.Quit ();

            //распределение коллекции по статьям с пометкой ненужных на удаление
            //List<string> delItemList = new List<string> ();
            foreach (ComplexItem c in cItemAll.LComplexItem)
            {
                foreach (ComplexItem c2 in cItemAll.LComplexItem)
                {
                    if (c2.ParentName == c.Name)
                    {
                        c.AddComplexItem ( c2 );
                        c2.del = "del";
                        //delItemList.Add ( c2.Name );
                    }
                }
            }
            // удаление ненужных статей
            for (int i = cItemAll.LComplexItem.Count - 1; i > 0; i--)
            {
                if (cItemAll.LComplexItem[i].del == "del")
                    //if (((ComplexItem)(cItemAll.LComplexItem[i])).del == "del") //изменено
                    cItemAll.LComplexItem.Remove ( cItemAll.LComplexItem[i] );
            }
            //foreach (string s in delItemList)
            //    cItemAll.LHashComplexItem.Remove ( s );
       //     wr.Close ();
            PrintItog ();


        }

        private void button2_Click ( object sender, RoutedEventArgs e )
        {
            PrintItog ();
        }
        void PrintItog ()
        {

            //   oExcel = new Excel.Application ();
            // = oExcel.Workbooks.Open ( path + "отчет.xls" );
            //  oExcel.Visible = true;
            oExcel.Application.ScreenUpdating = false;
            wb = oExcel.Workbooks.Add ( Type.Missing );

            ws = wb.Sheets[1];
            //ws.Cells[1, 1].value = "testtest";
            long frow = 5; // строка с которой начинаем

            ComplexItem.rLastRow = frow; //статический эл-т - последняя строка с которой работали
            foreach (ComplexItem c in cItemAll.LComplexItem) //в citemall - две статьи - доходы и расходы, для каждой из них вызывается метод вывода на лист и
            {
                c.PrintData ( ws );
                ComplexItem.rSeconLevelCount = 0; //wtf? сбрасывается счетчик чего?
            }
            ws.Range[ws.Cells[frow, 1], ws.Cells[ComplexItem.rLastRow, 7]].Borders.weight = 2;
            ws.Range[ws.Cells[frow, 3], ws.Cells[ComplexItem.rLastRow, 4]].NumberFormat = "#,##0.00";
            ws.Columns[2].EntireColumn.Autofit ();
            ws.Columns[3].EntireColumn.Autofit();
            ws.Columns[4].EntireColumn.Autofit();
            oExcel.DisplayAlerts = false;
            wb.SaveAs ( path + "DKKReport.xlsx" );
            oExcel.DisplayAlerts = true;
              wb.Close ();
              oExcel.Quit ();
          //  GetMarker ();
        }
        void GetMarker () //получение  цветов и комментов строки. Присваивание деляется строго по имени статьи, посему важно чтобы они все были уникальны
        {
            //  oExcel = new Excel.Application ();
            // oExcel.Visible = true;
            wbMark = oExcel.Workbooks.Open ( path + "marker.xlsm" ); //открываем файл с маркерами
    //        wsMark = wbMark.Worksheets.get_Item ( "marker" );
            wsMark = wbMark.Sheets["marker"];
            //    long ilMark = wsMark.UsedRange.Rows.Count;
            long ilMark = wsMark.Cells[ws.Rows.Count, 1].end ( Excel.XlDirection.xlUp ).row;
            long il = ws.Cells[ws.Rows.Count, 1].end ( Excel.XlDirection.xlUp ).row;
            //long  il=ws.Cells[ws.Rows.Count, 3].End ( Excel.XlDirection.xlUp ).row;
            long r = 0;
            wsMark.Columns[10].clear (); //стлбец с информацией - добавлена данная строка или нет
            for (int i = 1; i <= ilMark; i++) //цикл по всем строкам
            {
                r = 0;
                try
                {
                    r = ws.Range[ws.Cells[1, 2], ws.Cells[il, 2]].Find ( wsMark.Cells[i, 1].value ).row; //строки из фалй маркер ищется в файле  с данными
                }
                catch
                {
                }
                if (r > 0)
                {
                    wsMark.Cells[i, 10].value = "ok"; //если строка найдена в файл марекр ставится напроив нее ок
                    ws.Range[ws.Cells[r, 2], ws.Cells[r, 4]].Font.color = wsMark.Cells[i, 1].font.color; //переносится цвет
                    for (int j = 5; j < 8; j++) //переносятся комменты
                        ws.Cells[r, j].value = wsMark.Cells[i, j - 1].value;
                }
            } 
            oExcel.DisplayAlerts = false; //все сохраняется и закрывается
            wb.Save ();
            wbMark.Save ();

            oExcel.DisplayAlerts = true;
            wbMark.Close ();
            wb.Close ();
            oExcel.Application.ScreenUpdating = true;
            oExcel.Quit ();
        }
        void SaveMarker () //сохранение маркеров
        {
            File.Copy ( path + "marker.xlsm", path + "marker(old).xlsm", true ); //создается копия старого файла
            oExcel = new Excel.Application ();
            oExcel.Visible = true;

            wbMark = oExcel.Workbooks.Open ( path + "marker.xlsm" ); //открыается файл c маркерами
            //wsMark = wbMark.Worksheets.get_Item ( "marker" );
            wsMark = wbMark.Sheets["marker"];
            wb = oExcel.Workbooks.Open ( path + "DKKReport.xlsx" ); //открывается исходный файл
            ws = wb.Sheets[1];
            wsMark.UsedRange.Clear (); //марк файл очищается
            long il = ws.Cells[ws.Rows.Count, 1].end ( Excel.XlDirection.xlUp ).row;
            bool b;
            long k = 1;
            for (int i = 1; i < il; i++) //циклом по всем строкам ИСХОДНОГО файла
            {
                 b = false; //ключ - надо ли вообще сохранять эту строку

                for (int j = 5; j < 8; j++) //есть ли комменты (1)
                {
                    if (ws.Cells[i, j].value != null)
                        b = true;
                }
                if (ws.Cells[i, 2].font.color != 0) //есть ли фон (2)
                    b = true;
                if (b) //если что нить есть  - копируем всю строку и удаляем значения
                {
                    ws.Range[ws.Cells[i, 2], ws.Cells[i, 7]].Copy ( wsMark.Cells[k, 1] );
                    wsMark.Cells[k, 2].value = "";
                    wsMark.Cells[k++, 3].value = "";
                }

            }

            oExcel.DisplayAlerts = false;
            //wb.Save ();
            wbMark.Save ();
            oExcel.DisplayAlerts = true;
            wbMark.Close ();
            wb.Close ();
            oExcel.Application.ScreenUpdating = true;
            oExcel.Quit ();
        }

        private void button3_Click ( object sender, RoutedEventArgs e )
        {
            oExcel = new Excel.Application ();
            wb = oExcel.Workbooks.Open ( path + "DKKReport.xlsx" );
            ws = wb.Sheets[1];
            GetMarker ();
        }

        private void button4_Click ( object sender, RoutedEventArgs e )
        {
            SaveMarker ();
        }
    }
}
