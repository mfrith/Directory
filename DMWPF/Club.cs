using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using System.Data;
using System.Data.SqlClient;

namespace DMWPF
{
  class Club
  {
    public struct clubRecord
    {

      public int ClubNo;
      public string ClubName;
      public string Day;
      public string Time;
      public string Website;
      public string Phone;
      public string Email;
      public string Location;
      public string Address;
      public string City;
      public string Zip;
      public int Area;
      public string Division;
      public string Club_Type;
      public string Advanced;
      public string Facebook;
      private string _status;

      //public clubRecord(string[] rcd)
      //{
      //  ClubNo = System.Int32.Parse(rcd[2]);
      //  ClubName = rcd[3];
      //  Area = System.Int32.Parse(rcd[0]);
      //  Division = rcd[1];
      //  Location = rcd[4];
      //  Address = rcd[5];
      //  Day = rcd[10];
      //  Time = rcd[9];
      //  City = rcd[6];
      //  Zip = rcd[7];
      //  Club_Type = rcd[11];
      //  Advanced = rcd[15];
      //  Club_status = rcd[17];
      //  Phone = rcd[8];
      //  Email = rcd[13];
      //  Facebook = rcd[14];
      //  Website = rcd[12];
      //}
    }
  }


 
 
public class clsSqlCommandUpdate
{
    //Create Connection
    SqlConnection thisConnection = new SqlConnection("server=(local)\\SQLEXPRESS;" + "integrated security=sspi;database=Northwind");
 
    public void Main()
    {
 
        OpenConnection();
 
        //Insert Rows to make sure they exist
        Console.WriteLine("\n");
        Console.WriteLine("***Insert Rows to make sure they exist***");
        InsertRows();
 
        //Display Rows Before Update
        Console.WriteLine("\n");
        Console.WriteLine("***Display Rows Before Update***");
        SelectRows();
 
        //Update Rows
        Console.WriteLine("\n");
        Console.WriteLine("***Perform Update***");
        UpdateRows();
 
        //Display Rows after update
        Console.WriteLine("\n");
        Console.WriteLine("***Display Rows After Update***");
        SelectRows();
 
        //Clean up with delete of all inserted rows
        Console.WriteLine("\n");
        Console.WriteLine("***Clean Up By Deleting Inserted Rows***");
        DeleteRows();
 
        // Close Connection
        thisConnection.Close();
        Console.WriteLine("Connection Closed");
 
        Console.ReadLine();
    }

    void OpenConnection()
    {
        try
        {
            // Open Connection
            thisConnection.Open();
            Console.WriteLine("Connection Opened");
        }
        catch (SqlException ex)
        {
            // Display error
            Console.WriteLine("Error: " + ex.ToString());
        }
    }
    void SelectRows()
    {
 
        try
        {
            // Sql Select Query 
            string sql = "SELECT * FROM Employees";
            SqlCommand cmd = new SqlCommand(sql, thisConnection);
 
            SqlDataReader dr;
            dr = cmd.ExecuteReader();
            string strEmployeeID = "EmployeeID";
            string strFirstName = "FirstName";
            string strLastName = "LastName";
 
            Console.WriteLine("{0} | {1} | {2}", strEmployeeID.PadRight(10), strFirstName.PadRight(10), strLastName);
            Console.WriteLine("==========================================");
            while (dr.Read())
            {
                //reading from the datareader
                Console.WriteLine("{0} | {1} | {2}", 
                    dr["EmployeeID"].ToString().PadRight(10), 
                    dr["FirstName"].ToString().PadRight(10), 
                    dr["LastName"]);
            }
            dr.Close();
            Console.WriteLine("==========================================");
        }
 
        catch (SqlException ex)
        {
            // Display error
            Console.WriteLine("Error: " + ex.ToString());
        }
 
    }
 
    void InsertRows()
    {
 
 
        //Insert Rows to make sure row exists before updating
        //Create Command object
        SqlCommand nonqueryCommand = thisConnection.CreateCommand();
 
        try
        {
 
            // Create INSERT statement with named parameters
            nonqueryCommand.CommandText = "INSERT  INTO Employees (FirstName, LastName) VALUES (@FirstName, @LastName)";
 
            // Add Parameters to Command Parameters collection
            nonqueryCommand.Parameters.Add("@FirstName", SqlDbType.VarChar, 10);
            nonqueryCommand.Parameters.Add("@LastName", SqlDbType.VarChar, 20);
 
            // Prepare command for repeated execution
            nonqueryCommand.Prepare();
 
            // Data to be inserted
            string[] names = { "Wade", "David", "Charlie" };
            for (int i = 0; i < = 2; i++)
            {
                nonqueryCommand.Parameters["@FirstName"].Value = names[i];
                nonqueryCommand.Parameters["@LastName"].Value = names[i];
 
                Console.WriteLine("Executing {0}", nonqueryCommand.CommandText);
                Console.WriteLine("Number of rows affected : {0}", nonqueryCommand.ExecuteNonQuery());
            }
        }
        catch (SqlException ex)
        {
            // Display error
            Console.WriteLine("Error: " + ex.ToString());
        }
        finally
        {
 
 
        }
 
    }
 
    void UpdateRows()
    {
 
        try
        {
            // 1. Create Command
            // Sql Update Statement
            string updateSql = "UPDATE Employees " + "SET LastName = @LastName " + "WHERE FirstName = @FirstName";
            SqlCommand UpdateCmd = new SqlCommand(updateSql, thisConnection);
 
            // 2. Map Parameters
 
            UpdateCmd.Parameters.Add("@FirstName", SqlDbType.NVarChar, 10, "FirstName");
 
            UpdateCmd.Parameters.Add("@LastName", SqlDbType.NVarChar, 20, "LastName");
 
            UpdateCmd.Parameters["@FirstName"].Value = "Wade";
            UpdateCmd.Parameters["@LastName"].Value = "Harvey";
 
            UpdateCmd.ExecuteNonQuery();
        }
 
        catch (SqlException ex)
        {
            // Display error
            Console.WriteLine("Error: " + ex.ToString());
        }
 
    }
 
    void DeleteRows()
    {
 
        try
        {
            //Create Command objects
            SqlCommand scalarCommand = new SqlCommand("SELECT COUNT(*) FROM Employees", thisConnection);
 
            // Execute Scalar Query
            Console.WriteLine("Before Delete, Number of Employees = {0}", scalarCommand.ExecuteScalar());
 
 
            // Set up and execute DELETE Command
            //Create Command object
            SqlCommand nonqueryCommand = thisConnection.CreateCommand();
            nonqueryCommand.CommandText = "DELETE FROM Employees WHERE " + "Firstname='Wade'  or " + "Firstname='Charlie' AND Lastname='Charlie' or " + "Firstname='David' AND Lastname='David' ";
            Console.WriteLine("Executing {0}", nonqueryCommand.CommandText);
            Console.WriteLine("Number of rows affected : {0}", nonqueryCommand.ExecuteNonQuery());
 
            // Execute Scalar Query
            Console.WriteLine("After Delete, Number of Employee = {0}", scalarCommand.ExecuteScalar());
        }
 
        catch (SqlException ex)
        {
            // Display error
            Console.WriteLine("Error: " + ex.ToString());
        }
 
    }
 
}

}
