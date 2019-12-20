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
using MassuploadApplication;

namespace EmployeeManagement_ExcelData
{
    /// <summary>  
    /// Interaction logic for MainWindow.xaml  
    /// </summary>  
    public partial class MainWindow : Window
    {
        DataService _objExcelSer;
        Employee _emp = new Employee();

        public MainWindow()
        {
            InitializeComponent();
        }


        /// <summary>  
        /// Getting Data From Excel Sheet  
        /// </summary>  
        /// <param name="sender"></param>  
        /// <param name="e"></param>  
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            GetEmployeeData();
        }

        private void GetEmployeeData()
        {
            _objExcelSer = new DataService();
            try
            {
                dataGridEmployee.ItemsSource = _objExcelSer.ReadRecordFromEXCELAsync().Result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnRefreshRecord_Click(object sender, RoutedEventArgs e)
        {
            GetEmployeeData();
        }

        /// <summary>  
        /// Getting Data of each cell  
        /// </summary>  
        /// <param name="sender"></param>  
        /// <param name="e"></param>  
        private void dataGridEmployee_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            try
            {
                FrameworkElement emp_ID = dataGridEmployee.Columns[0].GetCellContent(e.Row);
                if (emp_ID.GetType() == typeof(TextBox))
                {
                    _emp.PSNO = Convert.ToInt32(((TextBox)emp_ID).Text);
                }

                FrameworkElement emp_Name = dataGridEmployee.Columns[1].GetCellContent(e.Row);
                if (emp_Name.GetType() == typeof(TextBox))
                {
                    _emp.Name = ((TextBox)emp_Name).Text;
                }

                FrameworkElement emp_Email = dataGridEmployee.Columns[2].GetCellContent(e.Row);
                if (emp_Email.GetType() == typeof(TextBox))
                {
                    _emp.Email = ((TextBox)emp_Email).Text;
                }

                

                FrameworkElement emp_Address = dataGridEmployee.Columns[4].GetCellContent(e.Row);
                if (emp_Address.GetType() == typeof(TextBox))
                {
                    _emp.Address = ((TextBox)emp_Address).Text;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>  
        /// Get entire Row  
        /// </summary>  
        /// <param name="sender"></param>  
        /// <param name="e"></param>  
        private void dataGridEmployee_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            try
            {
                bool IsSave = _objExcelSer.ManageExcelRecordsAsync(_emp).Result;
                if (IsSave)
                {
                    MessageBox.Show("Employee Record Saved Successfully.");
                }
                else
                {
                    MessageBox.Show("Some Problem Occured.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        /// <summary>  
        /// Get Record info to update  
        /// </summary>  
        /// <param name="sender"></param>  
        /// <param name="e"></param>  
        private void dataGridEmployee_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _emp = dataGridEmployee.SelectedItem as Employee;
        }
    }
}