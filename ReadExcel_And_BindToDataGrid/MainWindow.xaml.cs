using System;
using System.Windows;
using Microsoft.Win32;
using System.Data;
using System.Threading.Tasks;

namespace ReadExcel_And_BindToDataGrid
{
  
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            
        }




        private DataTable GenerateDataFromExcelFile(OpenFileDialog openFile)
        {
            
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(openFile.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(1); ;
                Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

                string strCellData = "";
                double douCellData;
                int row = 0;
                int col = 0;

                DataTable dt = new DataTable();
                for (col = 1; col <= excelRange.Columns.Count; col++)
                {
                    string strColumn = "";
                    strColumn = (string)(excelRange.Cells[2, col] as Microsoft.Office.Interop.Excel.Range).Value2;
                    dt.Columns.Add(strColumn, typeof(string));
                }


                for (row = 2; row <= excelRange.Rows.Count; row++)
                {

                    string strData = "";
                    for (col = 1; col <= excelRange.Columns.Count; col++)
                    {
                        try
                        {
                            strCellData = (string)(excelRange.Cells[row, col] as Microsoft.Office.Interop.Excel.Range).Value2;
                            strData += strCellData + "|";
                        }
                        catch (Exception ex)
                        {
                            douCellData = (excelRange.Cells[row, col] as Microsoft.Office.Interop.Excel.Range).Value2;
                            strData += douCellData.ToString() + "|";
                        }
                    }
                    strData = strData.Remove(strData.Length - 1, 1);
                    dt.Rows.Add(strData.Split('|'));

                }

                excelBook.Close(true, null, null);
                excelApp.Quit();


                return dt;

            }


        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
           
            {
                
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.DefaultExt = ".xlsx";
                openFile.Filter = "(.xlsx)|*.xlsx";

                
                var browseFile = openFile.ShowDialog();
                
                
                if (browseFile == true)
                {


                    
                    DataTable dt =GenerateDataFromExcelFile(openFile);
                    


                    // Sorts the datatable by first column 
                    dt.DefaultView.Sort = "Name ASC";
                    //Removing the dublicate row
                    dt.Rows.RemoveAt(0);
                    
                    dtGrid.ItemsSource = dt.DefaultView;
                    

                    excelImage.Visibility = Visibility.Hidden;
                    dtGrid.Visibility = Visibility.Visible;
                    


                }
               
            }
        }

    }
}
