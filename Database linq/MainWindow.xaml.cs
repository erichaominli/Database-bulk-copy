using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Database_linq
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void load_Click(object sender, RoutedEventArgs e)
        {
            // Create an instance of the open file dialog box.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog1.Filter = "Text Files (.xlsx)|*.xlsx|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;

            openFileDialog1.Multiselect = false;
            DialogResult result = openFileDialog1.ShowDialog();

            if (result.ToString() == "OK")
            {
                addbox.Text = openFileDialog1.FileName;
            }
            else
            {

            }
            var ep = new ExcelPackage(new FileInfo(addbox.Text));
            DataTable dt = ToDataTable(ep);
            dg.ItemsSource = dt.AsDataView();
        }
        public static DataTable ToDataTable(ExcelPackage package)
        {
            ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
            DataTable table = new DataTable();
            foreach (var firstRowCell in workSheet.Cells[1, 1, 1, workSheet.Dimension.End.Column])
            {
                table.Columns.Add(firstRowCell.Text);
            }
            for (var rowNumber = 2; rowNumber <= workSheet.Dimension.End.Row; rowNumber++)
            {
                var row = workSheet.Cells[rowNumber, 1, rowNumber, workSheet.Dimension.End.Column];
                var newRow = table.NewRow();
                foreach (var cell in row)
                {
                    newRow[cell.Start.Column - 1] = cell.Text;
                }
                table.Rows.Add(newRow);
            }

            //table.Columns.Add("isSend", typeof(string));
            return table;
        }

        private void import_Click(object sender, RoutedEventArgs e)
        {
            var ep = new ExcelPackage(new FileInfo(addbox.Text));
            DataTable dt = ToDataTable(ep);


            string connectionString = @"data source=tesql3;initial catalog=eric test;user id=web.user;password=webuser02182000";
            try
            {
                SqlConnection SqlConnectionObj = new SqlConnection(connectionString);
                SqlBulkCopy bulkCopy = new SqlBulkCopy(SqlConnectionObj, SqlBulkCopyOptions.TableLock | SqlBulkCopyOptions.FireTriggers | SqlBulkCopyOptions.UseInternalTransaction, null);
                bulkCopy.DestinationTableName = "[dbo].[ddaTest]";
                SqlConnectionObj.Open();
                bulkCopy.WriteToServer(dt);
                SqlConnectionObj.Close();
                System.Windows.MessageBox.Show("success");
            }
            catch (Exception ex)
            {
            }

            //DataTable table = new DataTable();
            //        SqlDataAdapter da = new SqlDataAdapter("Select * FROM dbo.ddaTest", connection);
            //        da.Fill(table);
            //        dg.ItemsSource = table.AsDataView();
            //        connection.Close();
                }


        private void clear_Click(object sender, RoutedEventArgs e)
        {
            string connectionString = @"data source=tesql3;initial catalog=eric test;user id=web.user;password=webuser02182000";
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand clear = new SqlCommand("Delete From [dbo].[ddaTest]", connection);
            clear.ExecuteNonQuery();
            connection.Close();
        }
    }
}
