using System;
using System.IO;
using System.Windows;
using Syncfusion.XlsIO;
using System.Linq;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace WpfApp1 {

    public partial class MainWindow : Window {

        private string Path;
        private string way;
        private string outFile;

        ExcelEngine excelEngine;
        IWorkbook workbook;
        IWorksheet sheet;
        IApplication application;
        FileStream stream;
        FileStream streamNew;

        ExcelEngine excelEngineNew;
        IWorkbook workbookNew;
        IWorksheet sheetNew;
        IApplication applicationNew;

        public MainWindow() {
            InitializeComponent();
        }

        private void CreateNameNewExcelFile() {
            outFile = string.Format("{0}{1}{2}{3}", way, "Результат-", DateTime.Now.ToString().Replace(":", ""), ".xls");
        }

        private void Button_Click(object sender, RoutedEventArgs e) {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "excel files |*.xls";
            if (openFileDialog.ShowDialog() == true) {
                txtEditor.Content = openFileDialog.FileName;
                Path = openFileDialog.FileName;
                way = openFileDialog.FileName.Remove(openFileDialog.FileName.IndexOf(openFileDialog.SafeFileName));
            }
        }

        private void CreateExcelFile() {
            excelEngineNew = new ExcelEngine();
            applicationNew = excelEngine.Excel;
            applicationNew.DefaultVersion = ExcelVersion.Excel2013;

            streamNew = new FileStream(outFile, FileMode.CreateNew, FileAccess.ReadWrite);
            workbookNew = applicationNew.Workbooks.Create(1);
            sheetNew = workbookNew.Worksheets[0];
        }

        private void Button_Click_1(object sender, RoutedEventArgs e) {

            if (!String.IsNullOrEmpty(Path)) {
                try {
                    OpenExcelFile();
                } catch {
                    MessageBox.Show("Файл открыт в другой программе. Закройте файл и повторите операцию.");
                    return;
                }

                if (CheckExcel()) {
                    CreateNameNewExcelFile();
                    CreateExcelFile();
                    ConvertToAvicenna();
                    workbookNew.SaveAs(streamNew);
                    stream.Dispose();
                    streamNew.Dispose();
                };

            } else {
                MessageBox.Show("Выберите файл для преобразования");
            }
        }

        private void ConvertToAvicenna()
        {
            int maxIndexHarmfuls = 0;
            for (int i = 2; i < sheet.Columns[0].LastRow + 1; i++) {
                // "ФИО"
                string[] FIO = ConvertFullName(sheet.Range[i, 2].Value);
                sheetNew.Range[i, 1].Value = FIO[0];
                sheetNew.Range[i, 2].Value = FIO[1];
                sheetNew.Range[i, 3].Value = FIO[2];
                // Дата рождения
                sheetNew.Range[i, 4].Value = sheet.Range[i, 4].Value;
                // Подразделение
                sheetNew.Range[i, 5].Value = sheet.Range[i, 6].Value;
                // Должность
                sheetNew.Range[i, 6].Value = sheet.Range[i, 3].Value;
                // Предприятие
                sheetNew.Range[i, 7].Value = sheet.Range[i, 5].Value;
                // Факторы
                string[] Harmfuls = ConvertHarmfulFactors(sheet.Range[i, 7].Value);
                int index = 0;

                foreach (string Harm in Harmfuls) {
                    sheetNew.Range[i, 8 + index].Text = Harm;
                    index++;
                    maxIndexHarmfuls = index > maxIndexHarmfuls ? index : maxIndexHarmfuls;
                }
            }

            // обзываем колонки 
            sheetNew.Range["A1"].Value = "Фамилия";
            sheetNew.Range["B1"].Value = "Имя";
            sheetNew.Range["C1"].Value = "Отчество";
            sheetNew.Range["D1"].Value = "Дата рождения";
            sheetNew.Range["E1"].Value = "Подразделение";
            sheetNew.Range["F1"].Value = "Должность";
            sheetNew.Range["G1"].Value = "Предприятие";

            // обзываем колонки факторов
            for (int i = 1; i < maxIndexHarmfuls + 1; i++) {
                sheetNew.Range[1, maxIndexHarmfuls + 4 + i].Value = "Вредность " + i;
            }
        }
        private void OpenExcelFile() {
            excelEngine = new ExcelEngine();
            application = excelEngine.Excel;
           // applicationNew.DefaultVersion = ExcelVersion.Excel2013;
            stream = new FileStream(Path, FileMode.Open, FileAccess.ReadWrite);
            workbook = application.Workbooks.Open(stream);
            sheet = workbook.Worksheets[0];
        }

        private bool CheckExcel() {
            string str = "@";
            string result;

            for (int i = 0; i < sheet.Rows[0].LastColumn; i++) {
                str += sheet.Rows[0].Cells[i].Value.ToString() + "@";
            }

            result = CheckColumn(str);

            if (result.Contains("true")) {
                return true;
            } else {
                MessageBox.Show("В файле отсутствует поле \"" + result + "\". \n Операция прервана.");
                return false;
            }
        }

        // проверяем, что все поля есть
        private string CheckColumn(string str) {
            string upStr = str.ToUpper();

            if (!upStr.Contains("@ФИО@")) return "ФИО";
            if (!upStr.Contains("@ДОЛЖНОСТЬ@")) return "ДОЛЖНОСТЬ";
            if (!upStr.Contains("@ДАТА РОЖДЕНИЯ@")) return "ДАТА РОЖДЕНИЯ";
            if (!upStr.Contains("@ПРЕДПРИЯТИЕ@")) return "ПРЕДПРИЯТИЕ";
            if (!upStr.Contains("@ПОДРАЗДЕЛЕНИЕ@")) return "ПОДРАЗДЕЛЕНИЕ";
            if (!upStr.Contains("@ФАКТОРЫ@")) return "ФАКТОРЫ";

            return "true";
        }

        // Методы преобразования
        private string[] ConvertFullName(string str) {
            string[] words = str.Split(' ');
            return words;
        }

        // строку в массив с факторами, для разбиения по ячейкам
        private string[] ConvertHarmfulFactors(string str) {
            string pattern = @"([А-я]\.)|([Пп]риложение.{0,3}[0-9]|[A-z]|[А-я,])";
            string[] result = Regex.Replace(str, pattern, String.Empty).Split(' ');
            return result.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
        }
    }
}