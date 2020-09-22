using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace ParseExcel
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public static Excel.Application excelapp = new Excel.Application();
            
        Workbooks excelappallworkbooks = excelapp.Workbooks;
        public static string CurDir = Environment.CurrentDirectory;
        //Открываем книгу и получаем на нее ссылку
        private static Workbook excelworkbook = excelapp.Workbooks.Open(CurDir + @"\Example.xlsx",
             Type.Missing, Type.Missing, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing);
        
        private static Sheets excelallsheets = excelworkbook.Worksheets;
        //Получаем ссылку на лист 1
        private static Worksheet excelworksheet = (Excel.Worksheet)excelallsheets.get_Item(1);
        //получаем номер последней ячейки
        private static Range lastcell = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
        int lastcellnumber = lastcell.Row;
        public int iter = 0;

        public MainWindow()
        {
            InitializeComponent();
        }

        public List<String> GetUniqValues(int columnNumber, Worksheet excelworksheet, int lastcellnumber)
        {
            List<String> items = new List<String>();
            items.Add(excelworksheet.get_Range("$A$2", Type.Missing).Value2);
            foreach (var item in excelworksheet.get_Range("$A$2" + ":A" + lastcellnumber, Type.Missing).Value2)
            {
                bool flag = false;
                foreach (var str in items )
                {
                    if (item.ToString() == str)
                    {
                        flag = true;
                        break;
                    }
                }
                if (!flag) items.Add(item.ToString());
            }
            return items;
        }

        private void CreateNewWorkbook(Worksheet excelworksheet, string unic_value, 
            Sheets excelallsheets, Workbooks excelappallworkbooks, string CurDir, int iterationNumber)
        {
            
            excelworksheet.get_Range("$A$1:$R$5781", Type.Missing)
                .AutoFilter(1, unic_value);

            var selectFilteredCells = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
            selectFilteredCells.Copy();
            var newexcelsheet = excelallsheets.Add().Paste();
            excelworksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            excelworksheet = (Excel.Worksheet)excelallsheets.get_Item(1);
            excelworksheet.Move();
            var misValue = Type.Missing;

            var newexcelbook = excelappallworkbooks.get_Item(2);
            newexcelbook.SaveAs(CurDir + @"\" + iterationNumber + @".xlsx", Type.Missing, Type.Missing, Type.Missing
                , Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges
                , Type.Missing, Type.Missing, Type.Missing, Type.Missing);      
            newexcelbook.Close();
        }

        private void buttonStart_Click(object sender, RoutedEventArgs e)
        {
            excelapp.Visible = true;
            excelworkbook.Activate();

            //выбираем первую ячейку по ней будем фильтровать
            var excelcells = excelworksheet.get_Range("A1", Type.Missing);

            List<string> UniqValues = GetUniqValues(1, excelworksheet, lastcellnumber);

            foreach (var item in UniqValues)
            {
                listBox1.Items.Add(item);
            }
        }

        private void filterInXlsxFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            iter += 1;
            CreateNewWorkbook(excelworksheet, listBox1.SelectedItem.ToString(), excelallsheets, excelappallworkbooks, CurDir, iter);
            rect_Status.Fill = System.Windows.Media.Brushes.Green;
        }

        private void listBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            rect_Status.Fill = System.Windows.Media.Brushes.Red;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Multiselect = true;
            openFile.DefaultExt = ".xlsx"; // Default file extension
            openFile.Filter = "Text documents (.xlsx)|*.xlsx"; // Filter files by extension
            openFile.InitialDirectory = CurDir;

            openFile.ShowDialog();

            Excel.Application excelapp1 = new Excel.Application();
            Workbooks excelappJoinallworkbooks = excelapp1.Workbooks;
            excelapp1.Visible = true;
            Workbook excelJoinXlsBook = excelapp1.Workbooks.Open(openFile.FileName,
         Type.Missing, Type.Missing, Type.Missing, Type.Missing,
         Type.Missing, Type.Missing, Type.Missing, Type.Missing,
         Type.Missing, Type.Missing, Type.Missing, Type.Missing,
         Type.Missing, Type.Missing);
            Sheets excelallsheetsJoin = excelJoinXlsBook.Worksheets;
            Worksheet excelworksheet = (Excel.Worksheet)excelallsheetsJoin.get_Item(1);
            Range lastcell = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int lastcellnumber = lastcell.Row;

            var newFileNames = openFile.FileNames.Skip(1);
            foreach(var item in newFileNames)
            {
                Workbook excelJoinXlsBook2 = excelapp1.Workbooks.Open(item,
         Type.Missing, Type.Missing, Type.Missing, Type.Missing,
         Type.Missing, Type.Missing, Type.Missing, Type.Missing,
         Type.Missing, Type.Missing, Type.Missing, Type.Missing,
         Type.Missing, Type.Missing);
                Sheets excelallsheetsJoin2 = excelJoinXlsBook2.Worksheets;
                Worksheet excelworksheet2 = (Excel.Worksheet)excelallsheetsJoin2.get_Item(1);
                Range lastcell2 = excelworksheet2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                int lastcellnumber2 = lastcell2.Row;
                var selectedCells2 = excelworksheet2.get_Range("$A$2:$R$" + lastcellnumber2, Type.Missing);
                selectedCells2.Copy();
                excelworksheet.Range["A2:R2"].Insert(XlInsertShiftDirection.xlShiftDown);


            }

            excelJoinXlsBook.SaveAs(CurDir + @"\Joined.xlsx", Type.Missing, Type.Missing, Type.Missing
                , Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges
                , Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelJoinXlsBook.Close();
        }
    }
}
