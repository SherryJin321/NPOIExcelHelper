using System;
using System.Collections.Generic;
using System.Linq;
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
using System.Data;
using ExcelHelper;

namespace NPOIExcelHelper
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        



        #region 导入Excel数据
        System.Data.DataTable importDataTable = new System.Data.DataTable();
        ExcelHelper.ExcelHelper importExcelHelper;

        private void ImportExcel_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.InitialDirectory = "d:\\";
            openFileDialog.Filter = "Microsoft Excel files(*.xls)|*.xls;*.xlsx";
            openFileDialog.FilterIndex = 1;

            bool? result = openFileDialog.ShowDialog();
            if (result == true)
            {
                importDataTable = new System.Data.DataTable();
                importExcelHelper = new ExcelHelper.ExcelHelper(openFileDialog.FileName);
                importDataTable = importExcelHelper.ExcelToDataTable("抽奖人员", true);

                dataGrid1.DataContext = importDataTable;
            }
        }

        #endregion


        #region 导出Excel数据
        System.Data.DataTable exportDataTable = new System.Data.DataTable();
        ExcelHelper.ExcelHelper exportExcelHelper;

        private void ExportExcel_Click(object sender, RoutedEventArgs e)
        {
            exportDataTable = importDataTable;
            exportExcelHelper = new ExcelHelper.ExcelHelper("d:\\导出名单.xlsx");
            exportExcelHelper.DataTableToExcel(exportDataTable, "导出名单", true);

            MessageBox.Show("导出成功");
        }

        #endregion
    }



}
