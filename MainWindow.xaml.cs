using DocumentFormat.OpenXml.ExtendedProperties;

using Microsoft.Win32;
using Microsoft.Win32;
using ExcelDataReader;
using System.Data;
using System.IO;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Wpf_Excel_to_Datagrid
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        IExcelDataReader edr;
        public MainWindow()
        {
            InitializeComponent();
        }
        // список для добавления dataSet причем только строк
        List<string> RowExcel = new List<string>();
        private void OpenExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "EXCEL Files (*.xlsx)|*.xlsx|EXCEL Files 2003 (*.xls)|*.xls|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() != true)
                return;

            DbGrig.ItemsSource = readFile(openFileDialog.FileName);
        }

        private DataView readFile(string fileNames)
        {
            // строка для устранения ошибки ::::
            //System.NotSupportedException: "No data is available for encoding 1252. For information on defining a custom encoding, see the documentation for the Encoding.RegisterProvider method."
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var extension = fileNames.Substring(fileNames.LastIndexOf('.'));
            try
            {
                // Создаем поток для чтения.
                FileStream stream = File.Open(fileNames, FileMode.Open, FileAccess.Read);
                // В зависимости от расширения файла Excel, создаем тот или иной читатель.
                // Читатель для файлов с расширением *.xlsx.
                if (extension == ".xlsx")
                    edr = ExcelReaderFactory.CreateOpenXmlReader(stream);
                // Читатель для файлов с расширением *.xls.
                else if (extension == ".xls")
                    edr = ExcelReaderFactory.CreateBinaryReader(stream);
                // После завершения чтения освобождаем ресурсы.
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\r" + "Возможно файл открыт" + "\n\r" + "Выходим из программы");
                // закрываем окно
                Win1.Close();
                // закрываем программу
                System.Environment.Exit(1);
            }
            //// reader.IsFirstRowAsColumnNames
            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };

            // Читаем, получаем DataView и работаем с ним как обычно.
            DataSet dataSet = edr.AsDataSet(conf);
            // выводим в datagrid
            DataView dtView = dataSet.Tables[0].AsDataView();

            // добавляем в список данные из таблицы excel
            foreach (DataTable dt in dataSet.Tables)
            {
                MessageBox.Show (dt.TableName); // название таблицы
                                                 // перебор всех столбцов
                foreach (DataColumn column in dt.Columns)
                    RowExcel.Add( column.ColumnName);

                // перебор всех строк таблицы  https://metanit.com/sharp/adonet/3.6.php
                foreach (DataRow row in dt.Rows)
                {
                    // получаем все ячейки строки
                    var cells = row.ItemArray;
                    foreach (object cell in cells)
                       MessageBox.Show (cell.ToString());
                }
            }
            edr.Close();
            return dtView;
        }
    }
}
