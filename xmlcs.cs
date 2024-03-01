using System.Windows;
using Microsoft.Win32;

namespace ExcelToSqlServer
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                txtFilePath.Text = openFileDialog.FileName;
            }
        }

        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            string excelFilePath = txtFilePath.Text;
            string connectionString = "Your SQL Server Connection String"; // Replace with your SQL Server connection string
            ExcelToSqlServerImporter importer = new ExcelToSqlServerImporter();
            importer.ImportDataFromExcel(excelFilePath, connectionString);
            MessageBox.Show("Import completed successfully!");
        }
    }
}
