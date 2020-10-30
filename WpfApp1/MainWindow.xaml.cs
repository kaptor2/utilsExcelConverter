using System;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;

namespace WpfApp1 {

    public partial class MainWindow : Window {

        private string Path;
        private string way;
        private string outFile;

        private HSSFWorkbook workbook;
        private HSSFSheet sheet;
        private Dictionary<string, int> fildsList;

        private void CreateNewExcelFile() {
            outFile = string.Format("{0}{1}{2}{3}", way, "Результат-", DateTime.Now.ToString().Replace(":", ""), ".xls");
            File.Copy(Path, outFile);
            // Console.WriteLine(string.Format("{0}{1}{2}{3}", way, "Результат-", DateTime.Now.ToString().Replace(":", ""), ".xls"));
        }

        public MainWindow() {
            InitializeComponent();
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

        private void Button_Click_1(object sender, RoutedEventArgs e) {

            fildsList = new Dictionary<string, int>();

            if (!String.IsNullOrEmpty(Path)) {
                try {
                    OpenExcelFile(Path);
                }
                catch {
                    MessageBox.Show("Файл открыт в другой программе. Закройте файл и повторите операцию.");
                    return;
                }
               
                if (CheckExcel()) {
                    CreateNewExcelFile();
                };
                // MessageBox.Show(sheet.GetRow(1).GetCell(1).ToString()); // пример чтения
            } else {
                MessageBox.Show("Выберите файл для преобразования");
            }
        }

        private void OpenExcelFile() {
            FileStream file = null;
            file = new FileStream(Path, FileMode.Open, FileAccess.Read);
            workbook = new HSSFWorkbook(file);
            sheet = (HSSFSheet)workbook.GetSheetAt(0);
        }

        private bool CheckExcel() {
            HSSFRow headerRow = (HSSFRow)sheet.GetRow(0);
            string str = "@";
            string result;

            for (int i = 0; i < headerRow.LastCellNum; i++) {
                fildsList.Add(headerRow.GetCell(i).ToString().ToUpper(), i);
                str += headerRow.GetCell(i).ToString() + "@";
            }

            result = CheckColumn(str);

            if (result.Contains("true")) {
                return true;
            } else {
                MessageBox.Show("В файле отсутствует поле \"" + result + "\". \n Операция прервана.");
                return false;
            }
        }

        private string CheckColumn(string str) {
            string upStr  = str.ToUpper();

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