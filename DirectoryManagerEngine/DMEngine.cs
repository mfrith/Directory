using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data.SqlClient;
using System.Data;

namespace DistrictManagerEngine
{

  public class DMLoader
  {
    SqlConnection conn = new SqlConnection();

    public struct clubRecord
    {
      //public int ClubID;
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

      public clubRecord(string[] rcd)
      {
        ClubNo = System.Int32.Parse(rcd[2]);
        ClubName = rcd[3];
        Area = System.Int32.Parse(rcd[0]);
        Division = rcd[1];
        Location = rcd[4];
        Address = rcd[5];
        Day = rcd[10];
        Time = rcd[9];
        City = rcd[6];
        Zip = rcd[7];
        Club_Type = rcd[11];
        Advanced = rcd[15];
        //Club_status = rcd[19];
        Phone = rcd[8];
        //Contact2 = rcd[21];
        //Fax = rcd[22];
        Email = rcd[13];
        Facebook = rcd[14];
        Website = rcd[12];
      }
    }

    public struct memberRecord
    {
      //public string District;
      //public string Division;
      //public int Area;
      //public string ClubName;
      public int ClubNumber;
      //public string JoinDate;
      //public string TermBeginDate;
      //public string TermEndDate;
      //public string MbrshipFulfillStatus;
      public int MemberID;
      public string LastName;
      public string FirstName;
      public string MiddleName;
      public string Title;
      //public string Gender;
      //public string Address1;
      //public string Address2;
      //public string Address3;
      //public string City;
      //public string State;
      //public string Zip;
      //public string Country;
      public string WorkPhone;
      public string HomePhone;
      //public string FaxPhone;
      public string CellPhone;
      public string Email;
      //public string Email2;
      //public string Web;

      public memberRecord(string[] rcd)
      {
        //District = rcd[0];
        //Division = rcd[1];
        //Area = System.Int32.Parse(rcd[2]);
        //ClubName = rcd[0];
        ClubNumber = System.Int32.Parse(rcd[0]);
        //JoinDate = rcd[5];
        //TermBeginDate = rcd[6];
        //TermEndDate = rcd[7];
        //MbrshipFulfillStatus = rcd[8];
        MemberID = System.Int32.Parse(rcd[1]);
        LastName = rcd[2];
        FirstName = rcd[3];
        MiddleName = rcd[4];
        Title = rcd[5];
        //Gender = rcd[14];
        //Address1 = rcd[6];
        //Address2 = rcd[7];
        //Address3 = rcd[8];
        //City = rcd[9];
        //State = rcd[10];
        //Zip = rcd[11];
        //Country = rcd[12];
        WorkPhone = rcd[6];
        HomePhone = rcd[7];
        CellPhone = rcd[8];
        //FaxPhone = rcd[9];
        Email = rcd[9];
        //Email2 = rcd[18];
        //Web = rcd[18];
      }
    }

    public struct officerRecord
    {
      //public string District;
      //public string Division;
      //public int Area;
      //public string ClubName;
      public int ClubNumber;
      //public string MbrshipFulfillStatus;
      //public string JoinDate;
      //public string officerTerm;
      public string office;
      //public string TermBeginDate;
      //public string TermEndDate;
      public int MemberID;
      //public string LastName;
      //public string FirstName;
      //public string MiddleName;
      //public string Title;
      ////public string Gender;
      //public string MailStop;
      //public string Address1;
      //public string Address2;
      //public string City;
      //public string State;
      //public string Zip;
      ////public string Country;
      //public string WorkPhone;
      //public string HomePhone;
      //public string FaxPhone;
      //public string CellPhone;
      //public string Email;
      //public string Email2;
      //public string Web;

      public officerRecord(string[] rcd)
      {
        //District = rcd[0];
        //Division = rcd[0];
        //Area = System.Int32.Parse(rcd[1]);
        //ClubName = rcd[2];
        ClubNumber = System.Int32.Parse(rcd[0]);
        //MbrshipFulfillStatus = rcd[5];
        //JoinDate = rcd[5];
        //officerTerm = rcd[6];
        office = rcd[1];
        //TermBeginDate = rcd[8];
        //TermEndDate = rcd[9];
        MemberID = System.Int32.Parse(rcd[2]);
        //LastName = rcd[6];
        //FirstName = rcd[7];
        //MiddleName = rcd[8];
        //Title = rcd[9];
        ////Gender = rcd[15];
        //MailStop = rcd[10];
        //Address1 = rcd[11];
        //Address2 = rcd[12];
        //City = rcd[13];
        //State = rcd[14];
        //Zip = rcd[15];
        ////Country = rcd[22];
        //WorkPhone = rcd[16];
        //HomePhone = rcd[17];
        //FaxPhone = rcd[18];
        //CellPhone = rcd[19];
        //Email = rcd[20];
        //Email2 = rcd[21];
        //Web = rcd[22];
      }
    }

    public void LoadOfficers()
    {
      conn.ConnectionString = @"Server=.\SQLEXPRESS;Database=D12;Integrated Security=true;";
      //FileStream fleReader = new FileStream("D:\\TI\\Databases\\July09\\officers.txt", FileMode.Open, FileAccess.Read);
      FileStream fleReader = new FileStream("G:\\TI\\2016July\\Officers.txt", FileMode.Open, FileAccess.Read);

      StreamReader stmReader = new StreamReader(fleReader);

      string line; //= stmReader.ReadLine(); //skip header

      //List<int> memberIDList = new List<int>();
      SqlCommand dbcmd = new SqlCommand();
      dbcmd.Connection = conn;
      conn.Open();

      CreateClubOfficersTable();
      //CreateUpdatedMembersTable();

      char[] delims = new char[] { '\t' };
      while ((line = stmReader.ReadLine()) != null)
      {
        string[] pole = line.Split(delims, StringSplitOptions.None);
        officerRecord rcd = new officerRecord(pole);
        string insertClubOfficer = "INSERT INTO ClubOfficers VALUES (" + rcd.ClubNumber + ",'" + rcd.office + "'," + rcd.MemberID + ")";
        dbcmd.CommandText = insertClubOfficer;
        dbcmd.ExecuteNonQuery();

        //if (memberIDList.Contains(rcd.MemberID))
        //  continue;
        //memberIDList.Add(rcd.MemberID);
        //string InsertMember = "INSERT INTO Members_08 VALUES (" + rcd.MemberID + ",'" + rcd.FirstName + "','" + rcd.MiddleName + "','" + rcd.LastName + "','" + rcd.Title + "' ,'" +
        //            rcd.MailStop + "','" + rcd.Address1 + "','" + rcd.Address2 + "','" + rcd.City + "','" + rcd.State + "','" + rcd.Zip + "','" + rcd.WorkPhone + "','" + rcd.HomePhone + "','" +
        //            rcd.CellPhone + "','" + rcd.Email + "','" + rcd.Email2 + "','" + rcd.Web + "')";

        //  Insert dept table records first
        //dbcmd.CommandText = InsertMember;
        //dbcmd.ExecuteNonQuery();
      }
    }

    public void LoadMembers()
    {
      conn.ConnectionString = @"Server=.\SQLEXPRESS;Database=D12;Integrated Security=true;";
      conn.Open();

      CreateMembersTable();
      CreateClubMembersTable();

      //FileStream fleReader = new FileStream("D:\\TI\\Databases\\July09\\members.txt", FileMode.Open, FileAccess.Read);
      FileStream fleReader = new FileStream("G:\\TI\\2016Jan\\members.txt", FileMode.Open, FileAccess.Read);

      StreamReader stmReader = new StreamReader(fleReader);

      string line;// = stmReader.ReadLine(); //skip header

      List<int> memberIDList = new List<int>();
      SqlCommand dbcmd = new SqlCommand();
      dbcmd.Connection = conn;
      //char[] delims = new char[] { ',', '\t' };
      char[] delims = new char[] { '\t' };
      while ((line = stmReader.ReadLine()) != null)
      {
        string[] pole = line.Split(delims, StringSplitOptions.None);
        memberRecord rcd = new memberRecord(pole);
        string insertClubMember = "INSERT INTO ClubMembers VALUES (" + rcd.ClubNumber + ",'" + rcd.MemberID + "')";
        dbcmd.CommandText = insertClubMember;
        dbcmd.ExecuteNonQuery();

        //int temps = 0;
        //if (rcd.MemberID == 409503)
        //  temps = 12;
        if (memberIDList.Contains(rcd.MemberID))
          continue;
        memberIDList.Add(rcd.MemberID);
        string InsertMember = "INSERT INTO Members VALUES (" + rcd.MemberID + ",'" + rcd.FirstName + "','" + rcd.MiddleName + "','" + rcd.LastName + "','" + rcd.Title + "' ,'" +
                    rcd.WorkPhone + "','" + rcd.HomePhone + "','" + rcd.CellPhone + "','" + rcd.Email + "')";
        //rcd.Address1 + "','" + rcd.Address2 + "','" + rcd.Address3 + "','" + rcd.City + "','" + rcd.State + "','" + rcd.Zip + "','" + rcd.Country + "','" + 
        //rcd.WorkPhone + "','" + rcd.HomePhone + "','" + rcd.CellPhone + "','" + rcd.Email + "')";


        //  Insert dept table records first
        dbcmd.CommandText = InsertMember;
        dbcmd.ExecuteNonQuery();
      }
    }

    public void CreateMembersTable()
    {
      string CreateMembersTable = "CREATE TABLE Members (MemberID INT PRIMARY KEY NOT NULL,"
      + "FirstName    VARCHAR(50),"
      + "MiddleName  varchar(50),"
      + "LastName varchar(50),"
      + "Title varchar(14),"
        //+ "Address1 varchar(50),"
        //+ "Address2 varchar(50),"
        //+ "Address3 varchar(50),"
        //+ "City varchar(40),"
        //+ "State varchar(3),"
        //+ "Zip varchar(15),"
        //+ "Country varchar(15),"
      + "WorkPhone varchar(40),"
      + "HomePhone varchar(25),"
      + "CellPhone varchar(30),"
      + "Email varchar(50))";
      //+ "Email2 varchar(40),"
      //+ "Web varchar(50))";

      SqlCommand DBCmd = new SqlCommand(CreateMembersTable, conn);
      //DBCmd.CommandText = CreateMembersTableSQL;
      DBCmd.ExecuteNonQuery();
    }

    public void CreateClubMembersTable()
    {
      string CreateClubMembersTable = "CREATE TABLE ClubMembers (ClubMemberID INT IDENTITY(1,1) PRIMARY KEY NOT NULL,"
      + "ClubNo int,"
      + "MemberID int)";

      SqlCommand dbCmd = new SqlCommand(CreateClubMembersTable, conn);
      dbCmd.ExecuteNonQuery();
    }

    public void CreateUpdatedMembersTable()
    {
      string CreateMembersTableSQL = "CREATE TABLE Members_08 (MemberID INT PRIMARY KEY NOT NULL,"
      + "FirstName    VARCHAR(50),"
      + "MiddleName  varchar(50),"
      + "LastName varchar(50),"
      + "Title char(10),"
      + "Mailstop varchar(50),"
      + "Address1 varchar(50),"
      + "Address2 varchar(50),"
      + "City varchar(40),"
      + "State varchar(3),"
      + "Zip varchar(10),"
      + "WorkPhone varchar(40),"
      + "HomePhone varchar(30),"
      + "CellPhone varchar(30),"
      + "Email varchar(40),"
      + "Email2 varchar(40),"
      + "Web varchar(50))";

      SqlCommand DBCmd = new SqlCommand(CreateMembersTableSQL, conn);
      //DBCmd.CommandText = CreateMembersTableSQL;
      DBCmd.ExecuteNonQuery();

      string CreateClubMembersTable = "CREATE TABLE ClubMembers_08 (ClubMemberID INT IDENTITY(1,1) PRIMARY KEY NOT NULL,"
      + "ClubNo int,"
      + "MemberID int)";

      SqlCommand dbCmd = new SqlCommand(CreateClubMembersTable, conn);
      dbCmd.ExecuteNonQuery();
    }

    public void CreateClubOfficersTable()
    {
      string CreateClubOfficersTableSQL = "CREATE TABLE ClubOfficers (ClubOfficerID INT IDENTITY(1,1) PRIMARY KEY NOT NULL,"
      + "ClubNo  int,"
      + "Office varchar(10),"
      + "MemberID int)";

      SqlCommand DBCmd = new SqlCommand(CreateClubOfficersTableSQL, conn);
      DBCmd.ExecuteNonQuery();
    }

    private bool TableExists(String tableName)
    {
      DataSet temp = new DataSet();
      SqlDataAdapter daT = new SqlDataAdapter("SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" + tableName + "]' AND xType = 'U')", conn);
      daT.Fill(temp);

      if (temp.Tables.Count < 0)
        return false;
      else
        return true;
    }

    private void DropTable(String tableName)
    {
      string dropTable = "DROP TABLE " + tableName;
      SqlCommand dbCmd = new SqlCommand(dropTable, conn);
      dbCmd.ExecuteNonQuery();

    }

    private void CreateClubsTable()
    {
      /*
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
       */

      string CreateClubTable = "CREATE TABLE Clubs (ClubID INT IDENTITY(1,1) NOT NULL,"
      + "ClubNo int PRIMARY KEY NOT NULL,"
      + "ClubName  varchar(max),"
      + "Division char(1),"
      + "Area int,"
      + "Location varchar(max),"
      + "Address varchar(max),"
      + "Day varchar(max),"
      + "Time varchar(max),"
      + "City varchar(max),"
      + "Zip char(10),"
      + "Type varchar(max),"
      + "Phone varchar(max),"
      + "Email varchar(max),"
      + "Facebook varchar(max),"
      + "Website varchar(max),"
      + "Advanced char(4))";

      SqlCommand dbCmd = new SqlCommand(CreateClubTable, conn);
      dbCmd.ExecuteNonQuery();
    }

    public void LoadClubs()
    {
      conn.ConnectionString = @"Server=.\SQLEXPRESS;Database=D12;Integrated Security=true;";
      conn.Open();
      FileStream fleReader = new FileStream("C:\\Users\\mike\\Documents\\TI\\clubs.txt", FileMode.Open, FileAccess.Read);

      StreamReader stmReader = new StreamReader(fleReader);

      string line;// = stmReader.ReadLine(); //skip header

      CreateClubsTable();

      char[] delims = new char[] { '\t' };
      while ((line = stmReader.ReadLine()) != null)
      {
        string[] pole = line.Split(delims, StringSplitOptions.None);
        string blah = pole[0];

        clubRecord rcd = new clubRecord(pole);

        string club = "INSERT INTO Clubs VALUES (" + rcd.ClubNo + ",'" + rcd.ClubName + "','" + rcd.Division
                     + "','" + rcd.Area + "','" + rcd.Location + "','" + rcd.Address + "','" + rcd.Day + "','" + rcd.Time + "','" + rcd.City + "','" + rcd.Zip
                     + "','" + rcd.Club_Type + "','" + rcd.Phone + "','" + rcd.Email + "','" + rcd.Facebook
                     + "','" + rcd.Website + "','" + rcd.Advanced + "')";
        //  Insert dept table records first
        SqlCommand dbcmd = new SqlCommand(club, conn);
        dbcmd.ExecuteNonQuery();
      }
    }

  }

  public class DMDirectoryBuilder
  {
  }

  public class Clubs
  {

  }
}
