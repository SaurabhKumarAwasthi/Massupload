using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using System.Data.OleDb;
using System.Data;
using System.IO;
namespace MassuploadApplication
{
   
        public class Employee
        {
            public int PSNO { get; set; }
            public string Name { get; set; }
            public string Email { get; set; }
            public string Contact { get; set; }
            public string Address { get; set; }
        }
    
    public class DataService
    {
        OleDbConnection Conn;
        OleDbCommand Cmd;

        public DataService()
        {
            string fileName = "Dataentry.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Excel", fileName);
            //string ExcelFilePath = @"\MassuploadApplication\Excel\Dataentry.xlsx";
            string excelConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;Persist Security Info=True";
            Conn = new OleDbConnection(excelConnectionString);
        }

        /// <summary>  
        /// Method to Get All the Records from Excel  
        /// </summary>  
        /// <returns></returns>  
        /// 
        
        public async Task<ObservableCollection<Employee>> ReadRecordFromEXCELAsync()
        {
            ObservableCollection<Employee> Employees = new ObservableCollection<Employee>();
            Conn.Open();
            Cmd = new OleDbCommand();
            Cmd.Connection = Conn;
            Cmd.CommandText = "Select * from [Sheet1$]";
            var Reader = await Cmd.ExecuteReaderAsync();
            while (Reader.Read())
            {
                Employees.Add(new Employee()
                {
                    PSNO = Convert.ToInt32(Reader["PSNO"]),
                    Name = Reader["Name"].ToString(),
                    Email = Reader["Email"].ToString(),
                    Contact = Reader["Contact"].ToString(),
                    Address = Reader["Address"].ToString()
                });
            }
            Reader.Close();
            Conn.Close();
            return Employees;
        }

        /// <summary>  
        /// Method to Insert Record in the Excel  
        /// S1. If the EmpNo =0, then the Operation is Skipped.  
        /// S2. If the Student is already exist, then it is taken for Update  
        /// </summary>  
        /// <param name="Emp"></param>  
        public async Task<bool> ManageExcelRecordsAsync(Employee emp)
        {
            bool IsSave = false;
            if (emp.PSNO != 0)
            {
                await Conn.OpenAsync();
                Cmd = new OleDbCommand();
                Cmd.Connection = Conn;
                Cmd.Parameters.AddWithValue("@PSNO", emp.PSNO);
                Cmd.Parameters.AddWithValue("@Name", emp.Name);
                Cmd.Parameters.AddWithValue("@Email", emp.Email);
                Cmd.Parameters.AddWithValue("@Contact", emp.Contact);
                Cmd.Parameters.AddWithValue("@Address", emp.Address);

                if (!IsStudentRecordExistAsync(emp).Result)
                {
                    Cmd.CommandText = "Insert into [Sheet1$] values (@PSNO,@Name,@Email,@Contact,@Address)";
                }
                else
                {
                    Cmd.CommandText = "Update [Sheet1$] set PSNO=@PSNO,Name=@Name,Email=@Email,Contact=@Contact,Address=@Address where PSNO=@PSNO";

                }
                int result = await Cmd.ExecuteNonQueryAsync();
                if (result > 0)
                {
                    IsSave = true;
                }
                Conn.Close();
            }
            return IsSave;

        }
        /// <summary>  
        /// The method to check if the record is already available   
        /// in the workgroup  
        /// </summary>  
        /// <param name="emp"></param>  
        /// <returns></returns>  
        private async Task<bool> IsStudentRecordExistAsync(Employee emp)
        {
            bool IsRecordExist = false;
            Cmd.CommandText = "Select * from [Sheet1$] where PSNO=@PSNO";
            var Reader = await Cmd.ExecuteReaderAsync();
            if (Reader.HasRows)
            {
                IsRecordExist = true;
            }

            Reader.Close();
            return IsRecordExist;
        }
    }
}  

