using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;
using System.Xml;


namespace DistrictManager
{

  public partial class Form1 : Form
  {
    SqlConnection conn = new SqlConnection();
    Word._Application oWord;
    Word._Document oDoc;
    object oMissing = System.Reflection.Missing.Value;
    object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

    string division = "";
    int area = 0;

    public Form1()
    {
      InitializeComponent();
    }

    private void toolStripMenuItem1_Click(object sender, EventArgs e)
    {
      /*  openToolStripMenuItem_Click(sender,  e);

            
        droptable_Click(sender, e);
             
        CreateTableBtn_Click(sender, e);

        FileStream fleReader = new FileStream("C:\\TI\\d12_members.csv", FileMode.Open, FileAccess.Read);
        StreamReader stmReader = new StreamReader(fleReader);

        string line = stmReader.ReadLine(); //skip header

        List<int> memberIDList  = new List<int>();
            
        char[] delims = new char[] { ',', '\t' };
        while ((line = stmReader.ReadLine()) != null)
        {
            string[] pole = line.Split(delims, StringSplitOptions.None);
            memberRecord rcd = new memberRecord(pole);
            //string insertClubMember = "INSERT INTO ClubMembers VALUES (" + rcd.ClubNumber + ",'" + rcd.MemberID + "')";
            //SqlCommand dbcmd = new SqlCommand(insertClubMember, conn);
            //dbcmd.ExecuteNonQuery();

            if (memberIDList.Contains(rcd.MemberID))
                continue;
            memberIDList.Add(rcd.MemberID);
            string InsertMember = "INSERT INTO Members VALUES (" + rcd.MemberID + ",'" + rcd.FirstName + "','" + rcd.MiddleName + "','" + rcd.LastName + "','" + rcd.Title + "' ,'" +
                        rcd.MailStop + "','" + rcd.Address1 + "','" + rcd.Address2 + "','" + rcd.City + "','" + rcd.State + "','" + rcd.Zip + "','" + rcd.WorkPhone + "','" + rcd.HomePhone + "','" +
                        rcd.CellPhone + "','" + rcd.Email + "','" + rcd.Email2 + "','" + rcd.Web + "')";
            //  Insert dept table records first
            SqlCommand dbcmd = new SqlCommand(InsertMember, conn);
            //dbcmd.CommandText = ;
            dbcmd.ExecuteNonQuery();

        }*/
      /*
      FileStream fleReader = new FileStream("C:\\TI\\d12officers.csv", FileMode.Open, FileAccess.Read);
      StreamReader stmReader = new StreamReader(fleReader);

      string line = stmReader.ReadLine(); //skip header

      char[] delims = new char[] { ',', '\t' };
      while ((line = stmReader.ReadLine()) != null)
      {
          string[] pole = line.Split(delims, StringSplitOptions.None);
          officerRecord rcd = new officerRecord(pole);
          string insertClubOfficer = "INSERT INTO ClubOfficers VALUES (" + rcd.ClubNumber + ",'" + rcd.office + "'," + rcd.MemberID + ")";
          try
          {
              SqlCommand dbcmd = new SqlCommand(insertClubOfficer, conn);
              int success = dbcmd.ExecuteNonQuery();
          }
          catch (Exception ex)
          {
              Console.WriteLine(ex.Message);
              return;
          }
      }
*/
      //AreaGovernorsTableAdapters.Area_GovernorsTableAdapter daAreaGov = new AreaGovernorsTableAdapters.Area_GovernorsTableAdapter();
      // AreaGovernors.Area_GovernorsDataTable datatableAreaGov = new AreaGovernors.Area_GovernorsDataTable();
      //daAreaGov.Fill(datatableAreaGov);
      //District12DataSetTableAdapters.D12_MembersTableAdapter da =  new DistrictManager.District12DataSetTableAdapters.D12_MembersTableAdapter();
      // District12DataSet.D12_MembersDataTable table = new District12DataSet.D12_MembersDataTable();
      //da.Fill(table);
      //District12DataSet.D12_MembersRow row = District12DataSet.D12_MembersRow(table.Rows[0]);

      //object name = row.LastName;
      //  object fname = row.FirstName;

      //District12DataSetTableAdapters.D12_ChairsTableAdapter chairs = new DistrictManager.District12DataSetTableAdapters.D12_ChairsTableAdapter();
      // District12DataSet.D12_ChairsDataTable chairsTable = new District12DataSet.D12_ChairsDataTable();

      //District12DataSet.D12_ChairsRow newRow = district12DataSet1.D12_Chairs.NewD12_ChairsRow();
      //DataRow ble = chairsTable.NewD12_ChairsRow();
      //newRow.Chair = "TLI";
      // newRow.MemberID = 24;
      //  district12DataSet1.D12_Chairs.Rows.Add(newRow);
      //d12_ChairsTableAdapter1.Update(district12DataSet1.D12_Chairs);
      /*
       */
      /*
      string InsertChairRecordsSQL = "INSERT INTO Chairs VALUES (12345,'Directory')";
      int memberid = 54344;
      string theChair = "TLIChair";
      string InsertChair = "INSERT INTO Chairs VALUES (" + memberid + ",'" + theChair + "')";
            
  //  Insert dept table records first
      try
      {
          SqlCommand DBCmd = new SqlCommand(InsertChairRecordsSQL, conn);
          DBCmd.ExecuteNonQuery();
      }
      catch (Exception ex)
      {
          Console.WriteLine(ex.Message);
          return;
      }

      SqlCommand dbcmd = new SqlCommand(InsertChair, conn);
      dbcmd.ExecuteNonQuery();
*/
      //SqlDataAdapter da = new SqlDataAdapter();
      //DataTable table = new DataTable();
      // DataSet ds = new DataSet();
      //  da.Fill(ds);
      //chairs.Insert("TLI", 23);
      //chairs.Update(
      //AreaGovernors.Area_GovernorsRow rowAreaGov in datatableAreaGov.Rows
      //Data
      //DataTable membersTable = new District12DataSet.D12_MembersDataTable(); 
      // DataRow newRow = table.NewRow();
      //District12DataSet.D12_MembersRow newRow = new District12DataSet.D12_MembersRow();
      /* 
       newRow[0] = 12345;
       newRow[1] = "1/1/1999";
       newRow[2] = "Frith";
       newRow[3] = "Michael";
       newRow[4] = "W";
       newRow[5] = "ATMG";
       newRow[6] = "";
       newRow[7] = "1604 Waterford Ave";
       newRow[8] = "";
       newRow[9] = "Redlands";
       newRow[10] = "CA";
       newRow[11] = "92374";
       newRow[12] = "909 793-2853";
       newRow[13] = "909 389-9678";
       newRow[14] = "";
       newRow[15] = "909 754-8626";
       newRow[16] = "mike@mednrey.com";
       newRow[17] = "";
       newRow[18] = "www.mednrey.com";
       */
      //da.Insert(123, "1/1/1888", "Frith", "Mike", "W", "ATMG", "", "", "1604 Waterford Ave", "Redlands", "CA",
      // "92374", "909793-2853", "909 389-9678", "", "909 754-8626", "mike@mednrey.com", "", "mednrey.com");

      //s.Close();
    }

    private void exitToolStripMenuItem_Click(object sender, EventArgs e)
    {
      conn.Close();
      this.Close();

    }

    private void CreateTableBtn_Click(object sender, EventArgs e)
    {
      // Create the tables
      /*             
      string CreateChairTableSQL = "CREATE TABLE Chairs"+
      "( ChairID  smallint IDENTITY(1,1) PRIMARY KEY NOT NULL,"+  
       "memberID int NOT NULL,"+
        "Chair        varchar(50)     NOT NULL)";

      string CreateAreaGovernorTable = "CREATE TABLE AreaGovernors ("
      + "AreaGovID INT IDENTITY(1,1) PRIMARY KEY NOT NULL,"
      + "Area varchar(2),"
      + "MemberID INT)";

      string CreateDivisionGovernorTable = "CREATE TABLE DivGovernors ("
      + "DivGovID INT IDENTITY(1,1) PRIMARY KEY NOT NULL,"
      + "Division varchar(1),"
      + "MemberID INT)";

      string CreateDistrictOfficersTable = "CREATE TABLE DistrictOfficers ("
      + "DistrictOfficerID INT IDENTITY(1,1) PRIMARY KEY NOT NULL,"
      + "Office varchar(30),"
      + "MemberID int)";

      try 
      {
          SqlCommand DBCmd = new SqlCommand(CreateChairTableSQL, conn);
      DBCmd.CommandText = CreateChairTableSQL;
      DBCmd.ExecuteNonQuery();
      }
      catch (Exception ex)
      {
      //Create tables failed
      Console.WriteLine (ex.Message);
      return;
      }

      // Now insert the records

      string[] InsertEmpRecordsSQL = {
           "insert into emp values (1,'JOHNSON','ADMIN',6,'12-17-1990',18000,NULL,4)",
           "insert into emp values (2,'HARDING','MANAGER',9,'02-02-1998',52000,300,3)",
           "insert into emp values (3,'TAFT','SALES I',2,'01-02-1996',25000,500,3)",
           "insert into emp values (4,'HOOVER','SALES I',2,'04-02-1990',27000,NULL,3)",
           "insert into emp values (5,'LINCOLN','TECH',6,'06-23-1994',22500,1400,4)",
           "insert into emp values (6,'GARFIELD','MANAGER',9,'05-01-1993',54000,NULL,4)",
           "insert into emp values (7,'POLK','TECH',6,'09-22-1997',25000,NULL,4)",
           "insert into emp values (8,'GRANT','ENGINEER',10,'03-30-1997',32000,NULL,2)",
           "insert into emp values (9,'JACKSON','CEO',NULL,'01-01-1990',75000,NULL,4)",
      "insert into emp values (10,'FILLMORE','MANAGER',9,'08-09-1994',56000, NULL,2)",
           "insert into emp values (11,'ADAMS','ENGINEER',10,'03-15-1996',34000, NULL,2)",
           "insert into emp values (12,'WASHINGTON','ADMIN',6,'04-16-1998',18000,NULL,4)",
           "insert into emp values (13,'MONROE','ENGINEER',10,'12-03-2000',30000,NULL,2)",
           "insert into emp values (14,'ROOSEVELT','CPA',9,'10-12-1995',35000,NULL,1)"};

      string[] InsertDeptRecordsSQL = {
           "insert into dept values (1,'ACCOUNTING','ST LOUIS')",
           "insert into dept values  (2,'RESEARCH','NEW YORK')",
           "insert into dept  values (3,'SALES','ATLANTA')",
           "insert into dept  values (4, 'OPERATIONS','SEATTLE')"};

      //  Insert dept table records first
      for (int x = 0; x<InsertDeptRecordsSQL.Length; x++)
      {
      try 
      {
          SqlCommand DBCmd = new SqlCommand(InsertDeptRecordsSQL[x], conn);
      DBCmd.ExecuteNonQuery();
      }
      catch (Exception ex) 
      {
      Console.WriteLine (ex.Message);
      return;
      }
      }

      //  Now the emp table records
      for (int x = 0; x<InsertEmpRecordsSQL.Length; x++)
      {
      try 
      {
          SqlCommand DBCmd = new SqlCommand(InsertEmpRecordsSQL[x], conn);
      DBCmd.ExecuteNonQuery();
      }
      catch (Exception ex) 
      {
      Console.WriteLine (ex.Message);
      return;
      }
      }

      Console.WriteLine ("Tables created Successfully!");

      //  Close the connection
      conn.Close(); 

      SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM Chairs", conn);
      DataSet ds = new DataSet();
      da.Fill(ds);
      SqlCommandBuilder bldr = new SqlCommandBuilder(da);
      DataTable tbl = ds.Tables["Table"];

      //object[] rowVals = new object[3];
      //rowVals[0] = 1;
      //rowVals[1] = 23;
      //rowVals[2] = "TLI";
      //DataRow insertedRow = tbl.Rows.Add(rowVals);

      DataRow row = ds.Tables[0].Rows[0];
      row.Delete();

      //tbl.Rows[0].Delete();
      //tbl.Rows[1].Delete();
      da.Update(ds);*/
      //SqlCommand cmd = new SqlCommand("SELECT * FROM Chairs", conn);
      /*
      District12DataSet.D12_ChairsRow newRow = district12DataSet1.D12_Chairs.NewD12_ChairsRow();
      newRow.Chair = "TLI";
      newRow.MemberID = 22;
      district12DataSet1.D12_Chairs.Rows.Add(newRow);
      d12_ChairsTableAdapter1.Update(district12DataSet1.D12_Chairs);
       * */
    }

    private void Form1_Load(object sender, EventArgs e)
    {
      conn.ConnectionString = @"Server=.\SQLEXPRESS;Database=District12;Integrated Security=true;";
      conn.Open();
        /*
      listViewOfficer.View = View.Details;
      listViewOfficer.Columns.RemoveAt(0);
      listViewOfficer.Columns.RemoveAt(0);
      //listViewOfficer.CheckBoxes = true;

      //listViewOfficer.Activation = ItemActivation.TwoClick;
      listViewOfficer.Columns.Add("Office", 100, HorizontalAlignment.Left);
      listViewOfficer.Columns.Add("First", 100, HorizontalAlignment.Left);
      listViewOfficer.Columns.Add("Last", 100, HorizontalAlignment.Left);
      listViewOfficer.Columns.Add("Title", 100, HorizontalAlignment.Left);
      listViewOfficer.Columns.Add("Member ID", 100, HorizontalAlignment.Left);
      listViewOfficer.Columns.Add("Address1", 100, HorizontalAlignment.Left);
      listViewOfficer.Columns.Add("Address2", 100, HorizontalAlignment.Left);
      listViewOfficer.Columns.Add("City", 100, HorizontalAlignment.Left);
      listViewOfficer.Columns.Add("State", 50, HorizontalAlignment.Left);
      listViewOfficer.Columns.Add("Zip", 50, HorizontalAlignment.Left);
      listViewOfficer.Columns.Add("WkPhone", 50, HorizontalAlignment.Left);
      listViewOfficer.Columns.Add("HmPhone", 50, HorizontalAlignment.Left);
      listViewOfficer.Columns.Add("CellPhone", 50, HorizontalAlignment.Left);
      listViewOfficer.Columns.Add("Email", 50, HorizontalAlignment.Left);
      //listViewOfficer.Columns.Add("Zip", 50, HorizontalAlignment.Left);

     
      ListViewItem item = new ListViewItem();
      item.Text = "office";
      item.SubItems.Add("first");
      item.SubItems.Add("last");
      item.SubItems.Add("title");
      item.SubItems.Add("memberid");
      item.SubItems.Add("address");
      listViewOfficer.Items.Add(item);
                                               
      TreeNode districtNode = new TreeNode();
      districtNode.Text = "District 12";
      treeView1.Nodes.Add(districtNode);

      TreeNode officersNode = new TreeNode();
      officersNode.Text = "Officers";
      districtNode.Nodes.Add(officersNode);

      DataSet dsOfficers = new DataSet();
      SqlDataAdapter daOfficers = new SqlDataAdapter("SELECT DistrictOfficers.Office From DistrictOfficers", conn);

      daOfficers.Fill(dsOfficers);
      DataTable dtOfficers = dsOfficers.Tables["Table"];

      foreach (DataRow rowOfficer in dtOfficers.Rows)
      {
        TreeNode officerNode = new TreeNode();
        officerNode.Text = rowOfficer.ItemArray[0].ToString().Trim();
        officersNode.Nodes.Add(officerNode);
      }

      TreeNode chairsNode = new TreeNode();
      chairsNode.Text = "Chairs";
      districtNode.Nodes.Add(chairsNode);

      DataSet dsChairs = new DataSet();
      SqlDataAdapter daChairs = new SqlDataAdapter("SELECT Chairs.Chair From Chairs", conn);

      daChairs.Fill(dsChairs);
      DataTable dtChairs = dsChairs.Tables["Table"];

      foreach (DataRow rowChair in dtChairs.Rows)
      {
        TreeNode chairNode = new TreeNode();
        chairNode.Text = rowChair.ItemArray[0].ToString().Trim();
        chairsNode.Nodes.Add(chairNode);
      }

      // add clubs
      DataSet dsMatrix = new DataSet();
      SqlDataAdapter daMatrix = new SqlDataAdapter("Select * FROM DivAreaMatrix", conn);
      daMatrix.Fill(dsMatrix);
      DataTable dtMatrix = dsMatrix.Tables[0];

      DataSet dsClubs = new DataSet();
      SqlDataAdapter daClubs = new SqlDataAdapter("SELECT ClubNo, Area, Division FROM Clubs", conn);

      daClubs.Fill(dsClubs);
      DataTable dataTableClub = dsClubs.Tables["Table"];

      foreach (DataRow row in dtMatrix.Rows)
      {
        String Division = row.ItemArray[1].ToString();
        object count = row.ItemArray[2];
        TreeNode divNode = new TreeNode();

        int numAreas = System.Convert.ToInt32(count);
        divNode.Text = "Division " + Division;
        districtNode.Nodes.Add(divNode);

        for (int area = 1; area <= numAreas; area++)
        {
          string selectString = "Division = " + "'" + Division + "'" + " AND Area = " + area;
          DataRow[] clubs = dataTableClub.Select(selectString);

          TreeNode areaNode = new TreeNode();
          areaNode.Text = "Area " + Division + area.ToString();
          divNode.Nodes.Add(areaNode);

          foreach (DataRow club in clubs)
          {
            TreeNode clubNode = new TreeNode();
            string ClubNo = club.ItemArray[0].ToString().Trim();
            clubNode.Text = ClubNo;
            areaNode.Nodes.Add(clubNode);
          }

        }
      }
      listViewClub.Dock = DockStyle.Fill;
      listViewOfficer.Dock = DockStyle.Fill;
      
      string fileName = Application.LocalUserAppDataPath + @"\MainForm.txt";
      using( StreamReader reader = new StreamReader(fileName))
      {
        string loc = reader.ReadLine();
        string sze = reader.ReadLine();
        PointConverter pc = new PointConverter();
        if (loc != null)
          Point pt1 = (Point)pc.ConvertFromString(.
          this.Location = (Point)TypeDescriptor.GetConverter(typeof(Point)).ConvertFromString(loc);
        if (sze != null)
          this.ClientSize = (Size)TypeDescriptor.GetConverter(typeof(Point)).ConvertFromString(reader.ReadLine());
      }*/
    }

    private void droptable_Click(object sender, EventArgs e)
    {


    }

    private void openToolStripMenuItem_Click(object sender, EventArgs e)
    {
      //conn.ConnectionString = @"Server=.\SQLEXPRESS;Database=D12;Integrated Security=true;";

      conn.Open();
    }

    private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
    {

    }

    private void importClubsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      //conn.ConnectionString = @"Server=.\SQLEXPRESS;Database=D12;Integrated Security=true;";

      //conn.Open();

      FileStream fleReader = new FileStream("C:\\TI\\clubs.csv", FileMode.Open, FileAccess.Read);
      StreamReader stmReader = new StreamReader(fleReader);

      string line;// = stmReader.ReadLine(); //skip header

      char[] delims = new char[] { ',' };
      while ((line = stmReader.ReadLine()) != null)
      {
        string[] pole = line.Split(delims, StringSplitOptions.None);
        //clubRecord rcd = new clubRecord(pole);

        //string club = "INSERT INTO Clubs VALUES (" + rcd.ClubNo + ",'" + rcd.ClubName + "','" + rcd.Days + "','" + rcd.Time + "','" + rcd.Website + "' ,'" +
        //            rcd.Phone + "','" + rcd.Email + "','" + rcd.Location1 + "','" + rcd.Location2 + "','" + rcd.Address + "','" + rcd.City + "','" + rcd.Zip + "','" + rcd.Area + "','" + rcd.Division + "')";
        //  Insert dept table records first
        //SqlCommand dbcmd = new SqlCommand(club, conn);
        //dbcmd.ExecuteNonQuery();
      }
    }

    private void generateDirectoryToolStripMenuItem_Click(object sender, EventArgs e)
    {
      //GenerateDirectory();
    }

    private void GenerateDivision(string division)
    {
      // Division
      DataSet dsDivStaff = new DataSet();
      SqlDataAdapter daDivStaff = new SqlDataAdapter("select office, memberid, priority, email from DivisionStaff " +
        "where Division = " + "'" + division + "' and memberid > 0 order by priority, office", conn);

      daDivStaff.Fill(dsDivStaff);
      DataTable dtDivStaff = dsDivStaff.Tables["Table"];

      int staffCount = dtDivStaff.Rows.Count;
      int remainder = 0;

      System.Math.DivRem(staffCount, 2, out remainder);

      if (remainder > 0)
        staffCount++;
      int blocksize = 5;

      int halfStaffCount = staffCount / 2;
      int totalRows = (int)halfStaffCount * blocksize;
      if (staffCount == 1)
        totalRows = blocksize;

      DataSet dsAreaGov = new DataSet();
      SqlDataAdapter daAreaGov = new SqlDataAdapter("Select * from DivAreaMatrix " +
        "where Division = " + "'" + division + "'", conn);

      daAreaGov.Fill(dsAreaGov);
      DataTable datatableAreaGov = dsAreaGov.Tables["Table"];
      DataRow rowAreaGov = datatableAreaGov.Rows[0];
      int areas = (int)rowAreaGov["Area"];
      string website = (string)rowAreaGov["WebSite"];
      string divStatus = (string)rowAreaGov["Status"];

      bool hasStatus = false;

      if (divStatus.Contains("S") || divStatus.Contains("D") || divStatus.Contains("P"))
        hasStatus = true;

      Word.Table divGovTitle;
      int offset = 0;
      Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      if (hasStatus)
      {
        divGovTitle = oDoc.Tables.Add(wrdRng, 2, 1, ref oMissing, ref oMissing);
        offset = 1;
      }
      else
        divGovTitle = oDoc.Tables.Add(wrdRng, 1, 1, ref oMissing, ref oMissing);

      divGovTitle.Range.ParagraphFormat.SpaceAfter = 0;
      divGovTitle.Cell(1, 1).Range.Text = "Division " + division;
      //divGovTitle.Cell(2, 1).Range.Text = website;
      divGovTitle.Rows[1].Range.Font.Bold = 1;
      divGovTitle.Rows[1].Range.Font.Size = 14;
      //divGovTitle.Rows[2].Range.Font.Bold = 1;
      //divGovTitle.Rows[2].Range.Font.Size = 9;

      if (hasStatus)
      {
        string divStat = "2014-2015 ";
        switch (divStatus)
        {
          case "P":
            divStat += "President's" + " ";
            break;
          case "S":
            divStat += "Select" + " ";
            break;
        }
        divStat += "Distinguished Division";

        divGovTitle.Cell(2, 1).Range.Text = divStat;
        divGovTitle.Rows[2].Range.Font.Bold = 1;
        divGovTitle.Rows[2].Range.Font.Size = 9;
      }
      divGovTitle.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
      //divGovTitle.Select();
      //oDoc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = 0;
      //divGovTitle.AllowPageBreaks = false;


      // populate Division governor and staff info
      Word.Table divisionGovernorTable;
      wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      divisionGovernorTable = oDoc.Tables.Add(wrdRng, totalRows, 2, ref oMissing, ref oMissing);
      divisionGovernorTable.Range.ParagraphFormat.SpaceAfter = 0;
      divisionGovernorTable.BottomPadding = 0;

      divisionGovernorTable.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
      divisionGovernorTable.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
      divisionGovernorTable.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
      divisionGovernorTable.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
      divisionGovernorTable.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth100pt;

      if (!hasStatus)
        divisionGovernorTable.Rows[2].Range.Font.Size = 4;
      string office;
      String email = "";
      string name = "";
      string loc1 = "";
      string loc2 = "";
      string phone1 = "";
      String phone2 = "";
      int memberID = 0;
      int counter = 1;
      int column = 1;

      foreach (DataRow rowDivStaff in dtDivStaff.Rows)
      {
        // empty strings

        email = "";
        name = "";
        loc1 = "";
        loc2 = "";
        phone1 = "";
        phone2 = "";
        office = rowDivStaff.ItemArray[0].ToString().Trim();
        memberID = (int)rowDivStaff.ItemArray[1];
        email = rowDivStaff.ItemArray[3].ToString().Trim();

        String blank = "";
        if (memberID > 0)
        {
          if (office == "Division Director")
          {
            GenerateMemberInfo(memberID, ref name, ref loc1, ref loc2, ref phone1, ref phone2, ref blank, false, true, false);
            email = division + "Dir@d12toastmasters.org";
          }
          else
            GenerateMemberInfo(memberID, ref name, ref loc1, ref loc2, ref phone1, ref phone2, ref email, false, true, false);
        }

        //if (counter > halfStaffCount)
        //  counter = 1;

        //if (phone2.Length < 1 && email.Length > 1)
        //{
        //  phone2 = email;
        //  email = "";
        //}


        divisionGovernorTable.Cell(3 + offset + (blocksize * (counter - 1)), column).Range.Text = office;
        divisionGovernorTable.Cell(4 + offset + (blocksize * (counter - 1)), column).Range.Text = name;
        //divisionGovernorTable.Cell(5 + offset + (blocksize * (counter - 1)), column).Range.Text = loc1;
        //divisionGovernorTable.Cell(6 + offset + (blocksize * (counter - 1)), column).Range.Text = loc2;
        divisionGovernorTable.Cell(5 + offset + (blocksize * (counter - 1)), column).Range.Text = phone1;
        //divisionGovernorTable.Cell(6 + offset + (blocksize * (counter - 1)), column).Range.Text = phone2;
        divisionGovernorTable.Cell(6 + offset + (blocksize * (counter - 1)), column).Range.Text = email;
        divisionGovernorTable.Rows[3 + offset + (blocksize * (counter - 1))].Range.Font.Bold = 1;
        divisionGovernorTable.Rows[3 + offset + (blocksize * (counter - 1))].Range.Font.Size = 11;
        divisionGovernorTable.Rows[4 + offset + (blocksize * (counter - 1))].Range.Font.Size = 9;
        divisionGovernorTable.Rows[5 + offset + (blocksize * (counter - 1))].Range.Font.Size = 9;
        divisionGovernorTable.Rows[6 + offset + (blocksize * (counter - 1))].Range.Font.Size = 9;
        //divisionGovernorTable.Rows[7 + offset + (blocksize * (counter - 1))].Range.Font.Size = 9;
        //divisionGovernorTable.Rows[8 + offset + (blocksize * (counter - 1))].Range.Font.Size = 9;
        //divisionGovernorTable.Rows[9 + offset + (blocksize * (counter - 1))].Range.Font.Size = 9;
        divisionGovernorTable.Rows[7 + offset + (blocksize * (counter - 1))].Range.Font.Size = 4;
        if (column == 1)
          column = 2;
        else if (column == 2)
        {
          counter++;
          column = 1;
        }
      }

      // empty strings

      email = "";
      name = "";
      loc1 = "";
      loc2 = "";
      phone1 = "";
      phone2 = "";

      Word.Paragraph divGovBreak;
      object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      divGovBreak = oDoc.Content.Paragraphs.Add(ref oRng);
      divGovBreak.Range.Font.Size = 4;
      divGovBreak.Format.SpaceAfter = 0;

      //DataSet dsAreaGov = new DataSet();
      //SqlDataAdapter daAreaGov = new SqlDataAdapter("Select * from DivAreaMatrix " +
      //  "where Division = " + "'" + division + "'", conn);

      //daAreaGov.Fill(dsAreaGov);
      //DataTable datatableAreaGov = dsAreaGov.Tables["Table"];
      //DataRow rowAreaGov = datatableAreaGov.Rows[0];
      //int areas = (int)rowAreaGov["Area"];
      for (int index = 1; index <= areas; index++)
      {
        //GenerateArea(division, index);
        GenerateAreaNew(oDoc, division, index);
      }

      Word.Table tTable;

      int nTables = oDoc.Tables.Count;
      if (nTables < 1)
        return;

      for (int index = 1; index <= nTables; index++)
      {
        tTable = oDoc.Tables[index];
        tTable.Select();
        tTable.Rows.AllowBreakAcrossPages = 0;
        oDoc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = -1;

        //for (int rowIndex = 1; rowIndex < tTable.Rows.Count; rowIndex++)
        //  tTable.Rows[rowIndex].Range.ParagraphFormat.KeepWithNext = -1;
      }

    }

    private void membersToolStripMenuItem_Click(object sender, EventArgs e)
    {
      DistrictManagerEngine.DMLoader DML = new DistrictManagerEngine.DMLoader();
      DML.CreateMembersTable();

      //string CreateMembersTableSQL = "CREATE TABLE Members (MemberID INT PRIMARY KEY NOT NULL,"
      //+ "FirstName    VARCHAR(50),"
      //+ "MiddleName  varchar(50),"
      //+ "LastName varchar(50),"
      //+ "Title char(10),"
      //+ "Mailstop varchar(50),"
      //+ "Address1 varchar(50),"
      //+ "Address2 varchar(50),"
      //+ "City varchar(40),"
      //+ "State varchar(3),"
      //+ "Zip varchar(10),"
      //+ "WorkPhone varchar(25),"
      //+ "HomePhone varchar(20),"
      //+ "CellPhone varchar(20),"
      //+ "Email varchar(40),"
      //+ "Email2 varchar(40),"
      //+ "Web varchar(50))";

      //SqlCommand DBCmd = new SqlCommand(CreateMembersTableSQL, conn);
      ////DBCmd.CommandText = CreateMembersTableSQL;
      //DBCmd.ExecuteNonQuery();

    }

    private void clubOfficersToolStripMenuItem_Click(object sender, EventArgs e)
    {
      DistrictManagerEngine.DMLoader DML = new DistrictManagerEngine.DMLoader();
      DML.CreateClubOfficersTable();

      //string CreateClubOfficersTableSQL = "CREATE TABLE ClubOfficers (ClubOfficerID INT IDENTITY(1,1) PRIMARY KEY NOT NULL,"
      //+ "ClubNo  int,"
      //+ "Office varchar(10),"
      //+ "MemberID int)";

      //SqlCommand DBCmd = new SqlCommand(CreateClubOfficersTableSQL, conn);
      ////DBCmd.CommandText = CreateClubOfficersTableSQL;
      //DBCmd.ExecuteNonQuery();

    }

    private void clubMembersToolStripMenuItem_Click(object sender, EventArgs e)
    {
      DistrictManagerEngine.DMLoader DML = new DistrictManagerEngine.DMLoader();
      DML.CreateClubMembersTable();

      string CreateClubMembersTable = "CREATE TABLE ClubMembers (ClubMemberID INT IDENTITY(1,1) PRIMARY KEY NOT NULL,"
      + "ClubNo int,"
      + "MemberID int)";

      SqlCommand dbCmd = new SqlCommand(CreateClubMembersTable, conn);
      dbCmd.ExecuteNonQuery();

    }

    private void membersToolStripMenuItem1_Click(object sender, EventArgs e)
    {
      string DropMembers = "DROP TABLE Members";
      SqlCommand dbCmd = new SqlCommand(DropMembers, conn);
      dbCmd.ExecuteNonQuery();
    }

    private void clubOfficersToolStripMenuItem1_Click(object sender, EventArgs e)
    {
      string DropClubOfficers = "DROP TABLE ClubOfficers";
      SqlCommand dbCmd = new SqlCommand(DropClubOfficers, conn);
      dbCmd.ExecuteNonQuery();
    }

    private void clubMembersToolStripMenuItem1_Click(object sender, EventArgs e)
    {
      string DropClubMembers = "DROP TABLE ClubMembers";
      SqlCommand dbCmd = new SqlCommand(DropClubMembers, conn);
      dbCmd.ExecuteNonQuery();
    }

    private void othersToolStripMenuItem_Click(object sender, EventArgs e)
    {
      string CreateAreaGovernorTable = "CREATE TABLE AreaGovernor (AreaGovernorID INT IDENTITY(1,1) PRIMARY KEY NOT NULL,"
      + "Division varchar(1),"
      + "Area  int,"
      + "MemberID int)";

      SqlCommand dbCmd = new SqlCommand(CreateAreaGovernorTable, conn);
      dbCmd.ExecuteNonQuery();

      //string CreateAssistantAreaGovernorTable = "CREATE TABLE AssistantAreaGovernor (AssistantAreaGovernorID INT IDENTITY(1,1) PRIMARY KEY NOT NULL,"
      //+ "Division varchar(1),"
      //+ "Area  int,"
      //+ "MemberID int)";

      //dbCmd.CommandText = CreateAssistantAreaGovernorTable;
      //dbCmd.ExecuteNonQuery();

      string CreateChairsTable = "CREATE TABLE Chairs (ChairID  smallint IDENTITY(1,1) PRIMARY KEY NOT NULL,"
      + "memberID int NOT NULL,"
      + "Chair varchar(50) NOT NULL)";

      dbCmd.CommandText = CreateChairsTable;
      dbCmd.ExecuteNonQuery();

      string CreateDistrictOfficersTable = "CREATE TABLE DistrictOfficers (DistrictOfficers  smallint IDENTITY(1,1) PRIMARY KEY NOT NULL,"
      + "memberID int NOT NULL,"
      + "Office varchar(50) NOT NULL)";

      dbCmd.CommandText = CreateDistrictOfficersTable;
      dbCmd.ExecuteNonQuery();

      string CreateDivisionGovernorTable = "CREATE TABLE DivisionGovernor (DivisionGovernorID INT IDENTITY(1,1) PRIMARY KEY NOT NULL,"
      + "Division varchar(10),"
      + "MemberID int)";

      dbCmd.CommandText = CreateDivisionGovernorTable;
      dbCmd.ExecuteNonQuery();

      //string CreateClubTable = "CREATE TABLE Clubs (ClubID INT IDENTITY(1,1) NOT NULL,"
      //+ "ClubNo int PRIMARY KEY NOT NULL,"
      //+ "ClubName  varchar(max),"
      //+ "Days varchar(50),"
      //+ "Time char(10),"
      //+ "Website varchar(max),"
      //+ "Phone varchar(25),"
      //+ "Email varchar(50),"
      //+ "Location1 varchar(50),"
      //+ "Location2 varchar(50),"
      //+ "Address varchar(50),"
      //+ "City varchar(50),"
      //+ "Zip char(10),"
      //+ "Area char(10),"
      //+ "Division char(10))";

      //dbCmd.CommandText = CreateClubTable;
      //dbCmd.ExecuteNonQuery();
    }

    private void calendarToolStripMenuItem_Click(object sender, EventArgs e)
    {
      FileStream fleReader = new FileStream("C:\\TI\\D12Calendar2007-2008.txt", FileMode.Open, FileAccess.Read);
      StreamReader stmReader = new StreamReader(fleReader);

      string line;// = stmReader.ReadLine(); //skip header

      SqlCommand dbcmd = new SqlCommand();
      dbcmd.Connection = conn;
      char[] delims = new char[] { '\t' };
      while ((line = stmReader.ReadLine()) != null)
      {
        string[] pole = line.Split(delims, StringSplitOptions.None);
        string date = pole[0];
        string dow = pole[1];
        string desc = pole[2];

        string insertCalendar = "INSERT INTO Calendar VALUES ('" + date.Trim() + "','" + dow.Trim() + "','" + desc.Trim() + "')";
        dbcmd.CommandText = insertCalendar;
        dbcmd.ExecuteNonQuery();
      }
    }

    private void membersToolStripMenuItem2_Click(object sender, EventArgs e)
    {

      DistrictManagerEngine.DMLoader DML = new DistrictManagerEngine.DMLoader();
      DML.LoadMembers();
      /*
      FileStream fleReader = new FileStream("D:\\TI\\Databases\\Jan08\\members_jan08.csv", FileMode.Open, FileAccess.Read);
      StreamReader stmReader = new StreamReader(fleReader);

      string line;// = stmReader.ReadLine(); //skip header

      List<int> memberIDList = new List<int>();
      SqlCommand dbcmd = new SqlCommand();
      dbcmd.Connection = conn;
      char[] delims = new char[] { ',', '\t' };
      while ((line = stmReader.ReadLine()) != null)
      {
        string[] pole = line.Split(delims, StringSplitOptions.None);
        //memberRecord rcd = new memberRecord(pole);
        //string insertClubMember = "INSERT INTO ClubMembers_Jan08 VALUES (" + rcd.ClubNumber + ",'" + rcd.MemberID + "')";
        //dbcmd.CommandText = insertClubMember;
        dbcmd.ExecuteNonQuery();

        //if (memberIDList.Contains(rcd.MemberID))
        //  continue;
        //memberIDList.Add(rcd.MemberID);
        //string InsertMember = "INSERT INTO Members_Jan08 VALUES (" + rcd.MemberID + ",'" + rcd.FirstName + "','" + rcd.MiddleName + "','" + rcd.LastName + "','" + rcd.Title + "' ,'" +
         //  /         rcd.MailStop + "','" + rcd.Address1 + "','" + rcd.Address2 + "','" + rcd.City + "','" + rcd.State + "','" + rcd.Zip + "','" + rcd.WorkPhone + "','" + rcd.HomePhone + "','" +
         //           rcd.CellPhone + "','" + rcd.Email + "','" + rcd.Email2 + "','" + rcd.Web + "')";

        //  Insert dept table records first
        //dbcmd.CommandText = InsertMember;
        dbcmd.ExecuteNonQuery();
      }*/
    }

    private void officersToolStripMenuItem_Click(object sender, EventArgs e)
    {
      DistrictManagerEngine.DMLoader DML = new DistrictManagerEngine.DMLoader();
      DML.LoadOfficers();
      /*
      FileStream fleReader = new FileStream("D:\\TI\\Directory\\Jan_08\\officers_jan08.csv", FileMode.Open, FileAccess.Read);
      StreamReader stmReader = new StreamReader(fleReader);

      string line = stmReader.ReadLine(); //skip header

      SqlCommand dbcmd = new SqlCommand();
      dbcmd.Connection = conn;

      char[] delims = new char[] { ',', '\t' };
      while ((line = stmReader.ReadLine()) != null)
      {
        string[] pole = line.Split(delims, StringSplitOptions.None);
        //officerRecord rcd = new officerRecord(pole);
        //string insertClubOfficer = "INSERT INTO ClubOfficers_Update VALUES (" + rcd.ClubNumber + ",'" + rcd.office + "'," + rcd.MemberID + ")";
        //dbcmd.CommandText = insertClubOfficer;
        dbcmd.ExecuteNonQuery();
      }*/
    }

    private void districtOfficersToolStripMenuItem_Click(object sender, EventArgs e)
    {
      oWord = new Word.Application();
      SetUpDocument();
      GenerateDistrictOfficers();
    }

    private void GenerateDistrictOfficers()
    {
      Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

      DataSet dsDistOfficers = new DataSet();
      /*
      SqlDataAdapter daDistOfficers = new SqlDataAdapter("SELECT DistrictOfficers.Office, Members.FirstName, Members.MiddleName,  Members.LastName, " +
          "Members.Title, Members.Mailstop, Members.Address1,Members.Address2, Members.City, Members.Zip, Members.WorkPhone, Members.HomePhone, " +
          "Members.CellPhone, Members.Email " +
          " FROM    DistrictOfficers INNER JOIN  Members ON DistrictOfficers.MemberID = Members.MemberID", conn);
       * */
      SqlDataAdapter daDistOfficers = new SqlDataAdapter("Select * from DistrictOfficers", conn);

      daDistOfficers.Fill(dsDistOfficers);
      DataTable dtDistOfficers = dsDistOfficers.Tables["Table"];

      Word.Table districtOfficersTitle;
      wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      districtOfficersTitle = oDoc.Tables.Add(wrdRng, 1, 1, ref oMissing, ref oMissing);
      districtOfficersTitle.Range.ParagraphFormat.SpaceAfter = 0;
      districtOfficersTitle.Cell(1, 1).Range.Text = "District 12 Officers";
      districtOfficersTitle.Rows[1].Range.Font.Bold = 1;
      districtOfficersTitle.Rows[1].Range.Font.Size = 14;
      districtOfficersTitle.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
      //districtOfficersTitle.Select();
      //oDoc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = -1;

      object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      Word.Table districtOfficersTable;
      wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      districtOfficersTable = oDoc.Tables.Add(wrdRng, 48, 2, ref oMissing, ref oMissing);
      districtOfficersTable.Range.ParagraphFormat.SpaceAfter = 0;
      districtOfficersTable.BottomPadding = 0;
      int counter = 1;
      int column = 1;

      string office;
      string firstName;
      string middleName;
      string lastName;
      string title;
      string mailstop;
      string address1;
      string address2;
      string city;
      string zip;
      string wkphone;
      string hmphone;
      string cellphone;
      string email = "";
      string name;
      string loc1;
      string loc2;
      string phone1;
      string phone2;
      int memberID;

      int blocksize = 6;
      foreach (DataRow rowDistOfficers in dtDistOfficers.Rows)
      {


        name = "";
        loc1 = "";
        loc2 = "";
        phone1 = "";
        phone2 = "";
        office = rowDistOfficers.ItemArray[2].ToString().Trim();
        memberID = (int)rowDistOfficers.ItemArray[1];

        if (memberID > 0)
          GenerateMemberInfo(memberID, ref name, ref loc1, ref loc2, ref phone1, ref phone2, ref email, false, true, false);

        email = rowDistOfficers.ItemArray[3].ToString().Trim();
        if (counter > 6)
        {
          column = 2;
          counter = 1;
        }

        //if (office == "District Governor")
        //  email = "dg@tmdistrict12.org";
        //else if (office == "LGET")
        //  email = "lget@tmdistrict12.org";
        //else if (office == "LGM")
        //  email = "lgm@tmdistrict12.org";

        //if (phone2.Length < 1 && email.Length > 1)
        //{
        //  phone2 = email;
        //  email = "";
        //}

        districtOfficersTable.Rows[2 + (blocksize * (counter - 1))].Range.Font.Size = 8;
        districtOfficersTable.Cell(3 + (blocksize * (counter - 1)), column).Range.Text = office;
        districtOfficersTable.Cell(4 + (blocksize * (counter - 1)), column).Range.Text = name;
        //districtOfficersTable.Cell(5 + (8 * (counter - 1)), column).Range.Text = loc1;
        //districtOfficersTable.Cell(6 + (8 * (counter - 1)), column).Range.Text = loc2;
        districtOfficersTable.Cell(5 + (blocksize * (counter - 1)), column).Range.Text = phone1;
        //districtOfficersTable.Cell(6 + (blocksize * (counter - 1)), column).Range.Text = phone2;
        districtOfficersTable.Cell(6 + (blocksize * (counter - 1)), column).Range.Text = email;
        districtOfficersTable.Rows[3 + (blocksize * (counter - 1))].Range.Font.Bold = 1;
        districtOfficersTable.Rows[3 + (blocksize * (counter - 1))].Range.Font.Size = 11;
        districtOfficersTable.Rows[4 + (blocksize * (counter - 1))].Range.Font.Size = 9;
        districtOfficersTable.Rows[5 + (blocksize * (counter - 1))].Range.Font.Size = 9;
        districtOfficersTable.Rows[6 + (blocksize * (counter - 1))].Range.Font.Size = 9;
        districtOfficersTable.Rows[7 + (blocksize * (counter - 1))].Range.Font.Size = 9;
        //districtOfficersTable.Rows[8 + (8 * (counter - 1))].Range.Font.Size = 9;
        //districtOfficersTable.Rows[9 + (8 * (counter - 1))].Range.Font.Size = 9;
        counter++;

      }

      DataSet dsDivGov = new DataSet();
      /*
      SqlDataAdapter daDivGov = new SqlDataAdapter("SELECT DivisionGovernor.Division, Members.FirstName, Members.MiddleName,  Members.LastName, " +
          "Members.Title, Members.Mailstop, Members.Address1,Members.Address2, Members.City, Members.Zip, Members.WorkPhone, Members.HomePhone, " +
          "Members.CellPhone, Members.Email " +
          " FROM    DivisionGovernor INNER JOIN  Members ON DivisionGovernor.MemberID = Members.MemberID", conn);
      */
      SqlDataAdapter daDivGov = new SqlDataAdapter("select Division, Office, MemberID, email from DivisionStaff " +
          "where office = 'Division Director' order by Division", conn);

      daDivGov.Fill(dsDivGov);
      DataTable datatableDivGov = dsDivGov.Tables["Table"];
      counter = 2;
      column = 2;
      foreach (DataRow rowDivGov in datatableDivGov.Rows)
      {
        
        name = "";
        loc1 = "";
        loc2 = "";
        phone1 = "";
        phone2 = "";
        string division = rowDivGov.ItemArray[0].ToString().Trim();
        office = rowDivGov.ItemArray[1].ToString().Trim();
        memberID = (int)rowDivGov.ItemArray[2];
        
        if (memberID > 0)
          GenerateMemberInfo(memberID, ref name, ref loc1, ref loc2, ref phone1, ref phone2, ref email, false, true, false);

        email = rowDivGov.ItemArray[3].ToString().Trim();

        if (phone1.Length < 1)
        {
          phone1 = phone2;
          phone2 = email;
          email = "";
        }

        if (phone1.Length < 1)
        {
          phone1 = email;
          phone2 = email = "";
        }

        if (phone2.Length < 1)
        {
          phone2 = email;
          email = "";
        }

        districtOfficersTable.Rows[2 + (blocksize * (counter - 1))].Range.Font.Size = 8;
        districtOfficersTable.Cell(3 + (blocksize * (counter - 1)), column).Range.Text = "Division " + division + " Director";
        districtOfficersTable.Cell(4 + (blocksize * (counter - 1)), column).Range.Text = name;
        //districtOfficersTable.Cell(5 + (8 * (counter - 1)), column).Range.Text = loc1;
        //districtOfficersTable.Cell(6 + (8 * (counter - 1)), column).Range.Text = loc2;
        districtOfficersTable.Cell(5 + (blocksize * (counter - 1)), column).Range.Text = phone1;
        districtOfficersTable.Cell(6 + (blocksize * (counter - 1)), column).Range.Text = phone2;
        districtOfficersTable.Cell(7 + (blocksize * (counter - 1)), column).Range.Text = email;
        districtOfficersTable.Rows[3 + (blocksize * (counter - 1))].Range.Font.Bold = 1;
        districtOfficersTable.Rows[3 + (blocksize * (counter - 1))].Range.Font.Size = 11;
        districtOfficersTable.Rows[4 + (blocksize * (counter - 1))].Range.Font.Size = 9;
        districtOfficersTable.Rows[5 + (blocksize * (counter - 1))].Range.Font.Size = 9;
        districtOfficersTable.Rows[6 + (blocksize * (counter - 1))].Range.Font.Size = 9;
        districtOfficersTable.Rows[7 + (blocksize * (counter - 1))].Range.Font.Size = 9;
        //districtOfficersTable.Rows[8 + (8 * (counter - 1))].Range.Font.Size = 9;
        //districtOfficersTable.Rows[9 + (8 * (counter - 1))].Range.Font.Size = 9;
        counter++;
      }

      Word.Paragraph paragraphBreak;
      oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      paragraphBreak = oDoc.Content.Paragraphs.Add(ref oRng);
      paragraphBreak.Format.SpaceAfter = 0;
    }

    private void SetUpDocument()
    {
      oWord.Visible = true;
      oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
      oDoc.PageSetup.LineNumbering.Active = 0;
      oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
      oDoc.PageSetup.TopMargin = oWord.InchesToPoints(.5F);
      oDoc.PageSetup.BottomMargin = 35;
      oDoc.PageSetup.LeftMargin = 35;
      oDoc.PageSetup.RightMargin = 25;
      oDoc.PageSetup.Gutter = 25;
      oDoc.PageSetup.HeaderDistance = 15;
      oDoc.PageSetup.FooterDistance = 15;
      //oDoc.PageSetup.PageWidth = 11;
      //oDoc.PageSetup.PageHeight = 8.5F;
      oDoc.PageSetup.FirstPageTray = Word.WdPaperTray.wdPrinterDefaultBin;
      oDoc.PageSetup.OtherPagesTray = Word.WdPaperTray.wdPrinterDefaultBin;
      oDoc.PageSetup.SectionStart = Word.WdSectionStart.wdSectionNewPage;
      oDoc.PageSetup.OddAndEvenPagesHeaderFooter = 0;
      oDoc.PageSetup.DifferentFirstPageHeaderFooter = 0;
      oDoc.PageSetup.VerticalAlignment = Word.WdVerticalAlignment.wdAlignVerticalTop;
      oDoc.PageSetup.SuppressEndnotes = 0;
      oDoc.PageSetup.MirrorMargins = 0;
      oDoc.PageSetup.TwoPagesOnOne = false;
      oDoc.PageSetup.BookFoldPrinting = true;
      //oDoc.PageSetup.BookFoldRevPrinting = true;
      //oDoc.PageSetup.BookFoldPrintingSheets = 1;
      oDoc.PageSetup.GutterPos = Word.WdGutterStyle.wdGutterPosLeft;

      // oDoc.Styles.
      // oDoc.Styles.Add
    }

    private void calendarToolStripMenuItem1_Click(object sender, EventArgs e)
    {
      oWord = new Word.Application();
      SetUpDocument();
      GenerateCalendar();
    }

    private void GenerateCalendar()
    {
      // Calendar  of Events
      //
      DataSet dsCalendar = new DataSet();
      SqlDataAdapter daCalendar = new SqlDataAdapter("SELECT * FROM Calendar", conn);
      daCalendar.Fill(dsCalendar);
      DataTable dtCalendar = dsCalendar.Tables["Table"];
      int rowCount = dtCalendar.Rows.Count;

      Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      Word.Table calendarTitle;
      wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      calendarTitle = oDoc.Tables.Add(wrdRng, 1, 1, ref oMissing, ref oMissing);
      calendarTitle.Range.ParagraphFormat.SpaceAfter = 0;
      calendarTitle.Cell(1, 1).Range.Text = "Calendar of Events";
      calendarTitle.Rows[1].Range.Font.Bold = 1;
      calendarTitle.Rows[1].Range.Font.Size = 14;
      calendarTitle.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
      //calendarTitle.Select();
      //oDoc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = -1;

      object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      Word.Table calendarTable;
      wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      calendarTable = oDoc.Tables.Add(wrdRng, rowCount, 3, ref oMissing, ref oMissing);
      calendarTable.Range.ParagraphFormat.SpaceAfter = 0;
      calendarTable.BottomPadding = 0;
      int counter = 2;
      foreach (DataRow rowCalendar in dtCalendar.Rows)
      {
        string dte = rowCalendar[1].ToString();
        string dow = rowCalendar[2].ToString();
        string desc = rowCalendar[3].ToString();
        calendarTable.Cell(counter, 1).Range.Text = dte.Trim();
        calendarTable.Cell(counter, 2).Range.Text = dow.Trim();
        calendarTable.Cell(counter, 3).Range.Text = desc.Trim();
        calendarTable.Rows[counter].Range.Font.Size = 9;
        counter++;
      }

      Word.Paragraph paragraphBreak;
      oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      paragraphBreak = oDoc.Content.Paragraphs.Add(ref oRng);
      paragraphBreak.Format.SpaceAfter = 0;
    }

    private void directoryToolStripMenuItem_Click(object sender, EventArgs e)
    {
      Thread dirThread = new Thread(new ThreadStart(GenerateDirectory));
      dirThread.Start();
    }

    private void GenerateDirectory()
    {
      oWord = new Word.Application();
      SetUpDocument();
      //GenerateDirectory();
      GenerateDistrictOfficers();
      GenerateChairs();

      DataSet dsDivisions = new DataSet();
      SqlDataAdapter daDiv = new SqlDataAdapter("Select * from DivAreaMatrix", conn);

      daDiv.Fill(dsDivisions);
      DataTable dtDivisions = dsDivisions.Tables["Table"];
      foreach (DataRow rowDivision in dtDivisions.Rows)
      {
        GenerateDivision(division);
      }
    }

    private void chairsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      Thread chairThread = new Thread(new ThreadStart(GenerateChairs));
      chairThread.Start();
    }

    public class chair : IComparable
    {
      public string theChair;
      public int MemberID;

      public chair(string inchair, int memberID)
      {
        theChair = inchair;
        MemberID = memberID;
      }

      public int CompareTo(object obj)
      {
        chair other = obj as chair;
        int result = this.theChair.CompareTo(other.theChair);
        return result;
      }
    }

    //public class chairSorter : IComparer<theChairs>
    //{

    //}

    private void GenerateChairs()
    {
      oWord = new Word.Application();
      SetUpDocument();
      //GenerateChairs();
      Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

      DataSet dsChairs = new DataSet();
      SqlDataAdapter daChairs = new SqlDataAdapter("select * from Chairs where memberID > 0", conn);
      /*
      SqlDataAdapter daChairs = new SqlDataAdapter("SELECT Chairs.Chair, Members.FirstName, Members.MiddleName,  Members.LastName, " +
          "Members.Title, Members.Mailstop, Members.Address1,Members.Address2, Members.City, Members.Zip, Members.WorkPhone, Members.HomePhone, " +
          "Members.CellPhone, Members.Email " +
          " FROM    Chairs INNER JOIN  Members ON Chairs.MemberID = Members.MemberID", conn);
       * */
      daChairs.Fill(dsChairs);
      DataTable dtChairs = dsChairs.Tables["Table"];

      Word.Table chairsTitle;
      wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      chairsTitle = oDoc.Tables.Add(wrdRng, 1, 1, ref oMissing, ref oMissing);
      chairsTitle.Range.ParagraphFormat.SpaceAfter = 0;
      chairsTitle.Cell(1, 1).Range.Text = "District 12 Chairs";
      chairsTitle.Rows[1].Range.Font.Bold = 1;
      chairsTitle.Rows[1].Range.Font.Size = 14;
      chairsTitle.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
      //chairsTitle.Select();
      //oDoc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = -1;

      int chairCount = dtChairs.Rows.Count;
      int zeroID = 0;
      int memberID;
      string chair = string.Empty;
      //chair[] chairs;
      List<chair> chairs = new List<chair>();

      foreach (DataRow rowChairs in dtChairs.Rows)
      {
        //sortedChairList.Add(rowChairs.ItemArray[1].ToString().Trim(), (int)rowChairs.ItemArray[2]);
        chair theChair = new chair(rowChairs.ItemArray[2].ToString().Trim(), (int)rowChairs.ItemArray[1]);
        if (theChair.MemberID > 0)
          chairs.Add(theChair);
        //chairs.SetValue(theChair);
      }
      int remainder = 0;

      System.Math.DivRem(chairCount, 2, out remainder);
      chairs.Sort();
      int chairDocRowCount = chairCount / 2;
      if (remainder > 0)
        chairDocRowCount++;

      //int totalRows = (int)chairDocRowCount * 8;
      int totalRows = (int)chairDocRowCount * 6;

      object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      Word.Table chairsTable;
      wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      chairsTable = oDoc.Tables.Add(wrdRng, totalRows, 2, ref oMissing, ref oMissing);
      chairsTable.Range.ParagraphFormat.SpaceAfter = 0;
      chairsTable.BottomPadding = 0;

      string office = string.Empty;
      string firstName = string.Empty;
      string middleName = string.Empty;
      string lastName = string.Empty;
      string title = string.Empty;
      string mailstop = string.Empty;
      string address1 = string.Empty;
      string address2 = string.Empty;
      string city = string.Empty;
      string zip = string.Empty;
      string wkphone = string.Empty;
      string hmphone = string.Empty;
      string cellphone = string.Empty;
      string email = string.Empty;
      string name = string.Empty;
      string loc1 = string.Empty;
      string loc2 = string.Empty;
      string phone1 = string.Empty;
      string phone2 = string.Empty;

      int counter = 0;
      int row = 0;
      //foreach (DataRow rowChairs in dtChairs.Rows)
      //foreach (KeyValuePair<string, int> chairMemberID in sortedChairList)
      //for (int index = 0; index < halfChairCount; index++)
      while (row < chairDocRowCount)
      {
        //if (counter >= chairCount)
        //  break;
        int column = 1;
        //int ID = 0;
        email = "";
        name = "";
        loc1 = "";
        loc2 = "";
        phone1 = "";
        phone2 = "";
        office = chairs[counter].theChair;
        memberID = chairs[counter].MemberID;
        if (memberID > 0)
          GenerateMemberInfo(memberID, ref name, ref loc1, ref loc2, ref phone1, ref phone2, ref email, false, true, false);

        //if (phone2.Length < 1 && email.Length > 1)
        //{
        //  phone2 = email;
        //  email = "";
        //}

        if (phone1.Length < 1)
        {
          phone1 = phone2;
          phone2 = email;
          email = "";
        }

        if (phone1.Length < 1)
        {
          phone1 = email;
          phone2 = email = "";
        }

        if (phone2.Length < 1)
        {
          phone2 = email;
          email = "";
        }
        //chairsTable.Cell(3 + (8 * (row)), column).Range.Text = office;
        //chairsTable.Cell(4 + (8 * (row)), column).Range.Text = name;
        //chairsTable.Cell(5 + (8 * (row)), column).Range.Text = loc1;
        //chairsTable.Cell(6 + (8 * (row)), column).Range.Text = loc2;
        //chairsTable.Cell(7 + (8 * (row)), column).Range.Text = phone1;
        //chairsTable.Cell(8 + (8 * (row)), column).Range.Text = phone2;
        //chairsTable.Cell(9 + (8 * (row)), column).Range.Text = email;
        chairsTable.Cell(3 + (6 * (row)), column).Range.Text = office;
        chairsTable.Cell(4 + (6 * (row)), column).Range.Text = name;
        chairsTable.Cell(5 + (6 * (row)), column).Range.Text = phone1;
        chairsTable.Cell(6 + (6 * (row)), column).Range.Text = phone2;
        chairsTable.Cell(7 + (6 * (row)), column).Range.Text = email;
        counter++;

        column = 2;
        email = "";
        name = "";
        loc1 = "";
        loc2 = "";
        phone1 = "";
        phone2 = "";
        if (counter >= chairs.Count)
        {
          chairsTable.Rows[2 + (6 * (row))].Range.Font.Size = 8;
          chairsTable.Rows[3 + (6 * (row))].Range.Font.Bold = 1;
          chairsTable.Rows[3 + (6 * (row))].Range.Font.Size = 9;
          chairsTable.Rows[4 + (6 * (row))].Range.Font.Size = 9;
          chairsTable.Rows[5 + (6 * (row))].Range.Font.Size = 9;
          chairsTable.Rows[6 + (6 * (row))].Range.Font.Size = 9;
          chairsTable.Rows[7 + (6 * (row))].Range.Font.Size = 9;
          //chairsTable.Rows[8 + (8 * (row))].Range.Font.Size = 9;
          //chairsTable.Rows[9 + (8 * (row))].Range.Font.Size = 9;

          break;
        }
        office = chairs[counter].theChair;
        memberID = chairs[counter].MemberID;

        if (memberID > 0)
          GenerateMemberInfo(memberID, ref name, ref loc1, ref loc2, ref phone1, ref phone2, ref email, false, true, false);

        if (phone2.Length < 1 && email.Length > 1)
        {
          phone2 = email;
          email = "";
        }

        chairsTable.Cell(3 + (6 * (row)), column).Range.Text = office;
        chairsTable.Cell(4 + (6 * (row)), column).Range.Text = name;
        ///chairsTable.Cell(5 + (8 * (row)), column).Range.Text = loc1;
        //chairsTable.Cell(6 + (8 * (row)), column).Range.Text = loc2;
        chairsTable.Cell(5 + (6 * (row)), column).Range.Text = phone1;
        chairsTable.Cell(6 + (6 * (row)), column).Range.Text = phone2;
        chairsTable.Cell(7 + (6 * (row)), column).Range.Text = email;

        chairsTable.Rows[2 + (6 * (row))].Range.Font.Size = 8;
        chairsTable.Rows[3 + (6 * (row))].Range.Font.Bold = 1;
        chairsTable.Rows[3 + (6 * (row))].Range.Font.Size = 9;
        chairsTable.Rows[4 + (6 * (row))].Range.Font.Size = 9;
        chairsTable.Rows[5 + (6 * (row))].Range.Font.Size = 9;
        chairsTable.Rows[6 + (6 * (row))].Range.Font.Size = 9;
        chairsTable.Rows[7 + (6 * (row))].Range.Font.Size = 9;
        //chairsTable.Rows[8 + (8 * (row))].Range.Font.Size = 9;
        //chairsTable.Rows[9 + (8 * (row))].Range.Font.Size = 9;
        counter++;
        row++;
      }

      //chairsTable.Rows[2 + (8 * (row))].Range.Font.Size = 8;
      //chairsTable.Rows[3 + (8 * (row))].Range.Font.Bold = 1;
      //chairsTable.Rows[3 + (8 * (row))].Range.Font.Size = 11;
      //chairsTable.Rows[4 + (8 * (row))].Range.Font.Size = 9;
      //chairsTable.Rows[5 + (8 * (row))].Range.Font.Size = 9;
      //chairsTable.Rows[6 + (8 * (row))].Range.Font.Size = 9;
      //chairsTable.Rows[7 + (8 * (row))].Range.Font.Size = 9;
      //chairsTable.Rows[8 + (8 * (row))].Range.Font.Size = 9;
      //chairsTable.Rows[9 + (8 * (row))].Range.Font.Size = 9;

      Word.Paragraph paragraphBreak;
      oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      paragraphBreak = oDoc.Content.Paragraphs.Add(ref oRng);
      paragraphBreak.Format.SpaceAfter = 0;

      Word.Table tTable;

      int nTables = oDoc.Tables.Count;
      if (nTables < 1)
        return;

      for (int index = 1; index <= nTables; index++)
      {
        tTable = oDoc.Tables[index];
        tTable.Select();
        tTable.Rows.AllowBreakAcrossPages = 0;
        oDoc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = -1;

        //for (int rowIndex = 1; rowIndex < tTable.Rows.Count; rowIndex++)
        //  tTable.Rows[rowIndex].Range.ParagraphFormat.KeepWithNext = -1;
      }
    }

    void GenerateDivisionThreadStart()
    {
      oWord = new Word.Application();
      SetUpDocument();

      GenerateDivision(division);
    }

    void GenerateAreaThreadStart()
    {
      oWord = new Word.Application();
      SetUpDocument();

      GenerateAreaNew(oDoc, division, area);
    }

    private void cToolStripMenuItem_Click(object sender, EventArgs e)
    {
      division = "C";
      Thread divThread = new Thread(new ThreadStart(GenerateDivisionThreadStart));
      divThread.Start();
    }

    private void aToolStripMenuItem_Click(object sender, EventArgs e)
    {
      division = "A";
      Thread divThread = new Thread(new ThreadStart(GenerateDivisionThreadStart));
      divThread.Start();
    }

    private void bToolStripMenuItem_Click(object sender, EventArgs e)
    {
      division = "B";
      Thread divThread = new Thread(new ThreadStart(GenerateDivisionThreadStart));
      divThread.Start();
    }

    private void dToolStripMenuItem_Click(object sender, EventArgs e)
    {
      division = "D";
      Thread divThread = new Thread(new ThreadStart(GenerateDivisionThreadStart));
      divThread.Start();
    }

    private void allToolStripMenuItem_Click(object sender, EventArgs e)
    {
      oWord = new Word.Application();
      SetUpDocument();
      DataSet dsDivisions = new DataSet();
      SqlDataAdapter daDiv = new SqlDataAdapter("Select * from DivAreaMatrix", conn);

      daDiv.Fill(dsDivisions);
      DataTable dtDivisions = dsDivisions.Tables["Table"];
      foreach (DataRow rowDivision in dtDivisions.Rows)
      {
        GenerateDivision(rowDivision.ItemArray[1].ToString());
      }
    }

    private void c6ToolStripMenuItem_Click(object sender, EventArgs e)
    {
      oWord = new Word.Application();
      SetUpDocument();
      GenerateArea("C", 6);
    }

    private void GenerateArea(string division, int area)
    {
      Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

      DataSet dsAreaGov = new DataSet();
      SqlDataAdapter daAreaGov = new SqlDataAdapter("select Area, MemberID from AreaGovernor " +
        "where Division = " + "'" + division + "'" + "AND AreaGovernor.Area = " + area, conn);
      int memberID = 0;

      daAreaGov.Fill(dsAreaGov);
      DataTable datatableAreaGov = dsAreaGov.Tables["Table"];
      DataRow rowAreaGov = null;
      if (datatableAreaGov.Rows.Count > 0)
      {
        rowAreaGov = datatableAreaGov.Rows[0];
        memberID = (int)rowAreaGov.ItemArray[1];
      }

      string office = "";

      string state = "CA";
      string name = "";
      string loc1 = "";
      string loc2 = "";
      string phone1 = "";
      string phone2 = "";
      string email = "";

      if (memberID > 0)
        GenerateMemberInfo(memberID, ref name, ref loc1, ref loc2, ref phone1, ref phone2, ref email, false, true, false);

      Word.Table areaGovTitle;
      wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      areaGovTitle = oDoc.Tables.Add(wrdRng, 1, 1, ref oMissing, ref oMissing);
      areaGovTitle.Range.ParagraphFormat.SpaceAfter = 0;
      areaGovTitle.Cell(1, 1).Range.Text = "Area " + division + area;
      areaGovTitle.Rows[1].Range.Font.Bold = 1;
      areaGovTitle.Rows[1].Range.Font.Size = 14;
      // areaGovTitle.Rows[2].Range.Font.Size = 8;
      areaGovTitle.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
      //areaGovTitle.Select();
      //oDoc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = 0;

      // populate Area governor and assistant info
      Word.Table areaGovernorTable;
      wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      areaGovernorTable = oDoc.Tables.Add(wrdRng, 7, 2, ref oMissing, ref oMissing);
      areaGovernorTable.Range.ParagraphFormat.SpaceAfter = 0;
      //clubLocationTable.AllowAutoFit = true;
      areaGovernorTable.BottomPadding = 0;
      //clubLocationTable.Rows[1].Height = 15;

      areaGovernorTable.Cell(2, 1).Range.Text = "Area Director";
      areaGovernorTable.Cell(2, 2).Range.Text = "Assistant Area Director";
      areaGovernorTable.Cell(3, 1).Range.Text = name;
      areaGovernorTable.Cell(4, 1).Range.Text = loc1;
      areaGovernorTable.Cell(5, 1).Range.Text = loc2;
      areaGovernorTable.Cell(6, 1).Range.Text = phone1;
      areaGovernorTable.Cell(7, 1).Range.Text = phone2;
      areaGovernorTable.Cell(8, 1).Range.Text = email;

      DataSet dsAssistantAreaGov = new DataSet();
      SqlDataAdapter daAssistantAreaGov = new SqlDataAdapter("select Area, MemberID from AssistantAreaGovernor " +
        "where Division = " + "'" + division + "'" + "AND Area = " + area, conn);

      daAssistantAreaGov.Fill(dsAssistantAreaGov);

      DataTable datatableAssistantAreaGov = dsAssistantAreaGov.Tables["Table"];
      DataRow rowAsstAreaGov = null;
      if (datatableAssistantAreaGov.Rows.Count > 0)
      {
        rowAsstAreaGov = datatableAssistantAreaGov.Rows[0];
        memberID = (int)rowAsstAreaGov.ItemArray[1];
      }
      // empty strings

      email = "";
      name = "";
      loc1 = "";
      loc2 = "";
      phone1 = "";
      phone2 = "";

      if (memberID > 0)
        GenerateMemberInfo(memberID, ref name, ref loc1, ref loc2, ref phone1, ref phone2, ref email, false, true, false);

      areaGovernorTable.Cell(3, 2).Range.Text = name;
      areaGovernorTable.Cell(4, 2).Range.Text = loc1;
      areaGovernorTable.Cell(5, 2).Range.Text = loc2;
      areaGovernorTable.Cell(6, 2).Range.Text = phone1;
      areaGovernorTable.Cell(7, 2).Range.Text = phone2;
      areaGovernorTable.Cell(8, 2).Range.Text = email;

      areaGovernorTable.Rows[2].Range.Font.Bold = 1;
      areaGovernorTable.Rows[2].Range.Font.Size = 11;
      areaGovernorTable.Rows[3].Range.Font.Size = 9;
      areaGovernorTable.Rows[4].Range.Font.Size = 9;
      areaGovernorTable.Rows[5].Range.Font.Size = 9;
      areaGovernorTable.Rows[6].Range.Font.Size = 9;
      areaGovernorTable.Rows[7].Range.Font.Size = 9;
      areaGovernorTable.Rows[8].Range.Font.Size = 9;

      datatableAssistantAreaGov.Clear();

      areaGovernorTable.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
      areaGovernorTable.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
      areaGovernorTable.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
      areaGovernorTable.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
      areaGovernorTable.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth100pt;

      Word.Paragraph areaGovBreakafter;
      object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      areaGovBreakafter = oDoc.Content.Paragraphs.Add(ref oRng);
      areaGovBreakafter.Range.Font.Size = 8;
      areaGovBreakafter.Format.SpaceAfter = 0;

      //areaGovernorTable.Select();
      //areaGovernorTable.AllowPageBreaks = false;

      // empty strings

      email = "";
      name = "";
      loc1 = "";
      loc2 = "";
      phone1 = "";
      phone2 = "";

      // Clubs in Area
      DataSet dsClubs = new DataSet();
      SqlDataAdapter daClubs = new SqlDataAdapter("SELECT * FROM Clubs " +
          " WHERE Division = " + "'" + division + "'" + " AND Area = " + "'" + area + "'" +
           " Order by clubno", conn);

      daClubs.Fill(dsClubs);
      DataTable dataTableClub = dsClubs.Tables["Table"];
      string clubNo;
      string clubName;
      string dayOfTheWeek;
      string time;
      string web;
      string address;
      string phone;
      string city;
      string zip;

      foreach (DataRow rowClub in dataTableClub.Rows)
      {
        //object clubNumber = rowClub.ItemArray[0];

        // populate club info
        // add a table for the club information
        clubNo = rowClub.ItemArray[1].ToString().Trim();
        bool bPrisonClub = false;

        if (division == "E")
        {
          if (clubNo == "7187" || clubNo == "8486" || clubNo == "771553" || clubNo == "814824" ||
              clubNo == "722657")
            bPrisonClub = true;
        }

        if (division == "D")
        {
          if (clubNo == "782516" || clubNo == "1071907" /*|| clubNo == "1159531"*/ || clubNo == "1167482")
            bPrisonClub = true;
        }

        clubName = rowClub.ItemArray[2].ToString().Trim();
        dayOfTheWeek = rowClub.ItemArray[3].ToString().Trim();
        time = rowClub.ItemArray[4].ToString().Trim();
        web = rowClub.ItemArray[5].ToString().Trim();
        phone = rowClub.ItemArray[6].ToString().Trim();
        email = rowClub.ItemArray[7].ToString().Trim();
        loc1 = rowClub.ItemArray[8].ToString().Trim();
        //loc2 = rowClub.ItemArray[9].ToString().Trim();
        address = rowClub.ItemArray[9].ToString().Trim();
        city = rowClub.ItemArray[10].ToString().Trim();
        zip = rowClub.ItemArray[11].ToString().Trim();

        Word.Table clubLocationTable;
        wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
        clubLocationTable = oDoc.Tables.Add(wrdRng, 4, 1, ref oMissing, ref oMissing);
        clubLocationTable.Range.ParagraphFormat.SpaceAfter = 0;
        //clubLocationTable.AllowAutoFit = true;
        clubLocationTable.BottomPadding = 0;

        // select table to keep from splitting across page breaks
        //int LastRow = clubLocationTable.Range.Rows.Count;
        //Word.Range MyRange = clubLocationTable.Range.Rows[0].Range;
        //MyRange.SetRange(0, LastRow);
        //MyRange = ActiveDocument.Tables(1).Range.Rows(2).Range
        //MyRange.SetRange Start:=MyRange.Start, _
        //End:=ActiveDocument.Tables(1).Range.Rows(LastRow).Range.End 
        //clubLocationTable.Range.SetRange
        //clubLocationTable.Select();
        //oDoc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = -1;

        clubLocationTable.AllowPageBreaks = false;
        //clubLocationTable.Selection.ParagraphFormat.KeepWithNext = true;
        //MyRange.ParagraphFormat.KeepWithNext = true;
        //areaGovTitle.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
        //clubLocationTable.Rows[1].Height = 15;
        string clubRow1;
        string clubRow2 = "";
        string clubRow3 = "";
        string clubRow4 = "";
        string clubRow5 = "";

        // need logic in club info, if something is empty we don't want a bunch of commas
        clubRow1 = "Club " + clubNo + " - " + clubName;
        //clubRow2 = dayOfTheWeek + " " + time + " " + web + " " + phone + " " + email;
        if (dayOfTheWeek.Length > 1)
          clubRow2 += dayOfTheWeek + " ";

        if (time.Length > 1)
          clubRow2 += time + " ";

        if (web.Length > 1)
          clubRow2 += web + " ";

        if (phone.Length > 1)
          clubRow2 += phone + " ";

        if (email.Length > 1)
          clubRow2 += email;

        //clubRow3 = loc1 + ", " + loc2 + ", " + address + ", " + city + ", " + zip;
        if (loc1.Length > 1)
          clubRow3 += loc1 + ", ";

        //if (loc2.Length > 1)
        //  clubRow3 += loc2 + ", ";

        if (address.Length > 1)
          clubRow3 += address + ", ";

        if (city.Length > 1)
          clubRow3 += city + ", ";

        if (zip.Length > 1)
          clubRow3 += zip;

        if (bPrisonClub)
        {
          if (clubNo == "7187" || clubNo == "8486" || clubNo == "771553" || clubNo == "814824" ||
              clubNo == "722657")
          {
            clubRow2 = "";
            if (dayOfTheWeek.Length > 1)
              clubRow2 += dayOfTheWeek + " ";

            if (time.Length > 1)
              clubRow2 += time + " ";

            clubRow3 = "Contact: Special Environment Clubs Chair, Randy Amelino, DTM";
            clubRow4 = "Email: chemdryall@gnww.net Phone: 951-258-6901";
            bPrisonClub = true;

            if (clubNo == "7187" || clubNo == "8486")
              clubRow2 += "California Institute for Women, 16756 Chino-Corona Rd, Corona, 92880-9508";
            else if (clubNo == "771553" || clubNo == "814824")
              clubRow2 += "California Rehab Center, Bldg 601 Men’s Education Dept, Norco, 92860";
            else if (clubNo == "722657")
              clubRow2 += "California Rehab Center, Bldg 601 Men’s Education Dept, Norco, 92860";
          }

          if (clubNo == "782516" || clubNo == "1071907" || /*clubNo == "1159531" || */clubNo == "1167482")
          {
            clubRow2 = "";
            if (dayOfTheWeek.Length > 1)
              clubRow2 += dayOfTheWeek + " ";

            if (time.Length > 1)
              clubRow2 += time + " ";

            clubRow3 = "";
            clubRow4 = "";
            clubRow5 = "";

            if (clubNo == "782516" || clubNo == "1071907")
              clubRow2 += "Chuckwalla Staff, Cindy Nepusz, 760.922.5300x5265";
            else if (clubNo == "1159531" || clubNo == "1167482")
              clubRow2 += "Tameka Roberson, 760.922.0680x5074";

          }
        }


        clubLocationTable.Cell(1, 1).Range.Text = clubRow1;
        clubLocationTable.Cell(2, 1).Range.Text = clubRow2;
        clubLocationTable.Cell(3, 1).Range.Text = clubRow3;
        clubLocationTable.Cell(4, 1).Range.Text = clubRow4;
        clubLocationTable.Rows[1].Range.Font.Bold = 1;
        clubLocationTable.Rows[1].Range.Font.Size = 11;
        clubLocationTable.Rows[2].Range.Font.Size = 9;
        clubLocationTable.Rows[3].Range.Font.Size = 9;
        clubLocationTable.Rows[4].Range.Font.Size = 9;

        // determine which officer is selected
        Word.Table officersTable;
        wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
        if (bPrisonClub)
        {
          officersTable = oDoc.Tables.Add(wrdRng, 5, 2, ref oMissing, ref oMissing);
          officersTable.Range.ParagraphFormat.SpaceAfter = 0;
          //oTable.AllowAutoFit = true;
          officersTable.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
          officersTable.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
          officersTable.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
          officersTable.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
          officersTable.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth100pt;
          officersTable.Cell(6, 1).Range.Text = "President";
          officersTable.Cell(6, 2).Range.Text = "Vice President Education";
          officersTable.Cell(8, 1).Range.Text = "Vice President Membership";
          officersTable.Cell(8, 2).Range.Text = "Vice President Public Relations";
          officersTable.Rows[6].Range.Font.Bold = 1;
          officersTable.Rows[8].Range.Font.Bold = 1;
          officersTable.Rows[6].Range.Font.Size = 11;
          officersTable.Rows[8].Range.Font.Size = 11;
          officersTable.Rows[5].Range.Font.Size = 8;
          officersTable.Rows[7].Range.Font.Size = 9;
          officersTable.Rows[9].Range.Font.Size = 9;
        }
        else
        {
          officersTable = oDoc.Tables.Add(wrdRng, 10, 2, ref oMissing, ref oMissing);
          officersTable.Range.ParagraphFormat.SpaceAfter = 0;
          //oTable.AllowAutoFit = true;
          officersTable.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
          officersTable.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
          officersTable.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
          officersTable.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
          officersTable.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth100pt;
          officersTable.Cell(5, 1).Range.Text = "President";
          officersTable.Cell(5, 2).Range.Text = "Vice President Education";
          officersTable.Cell(11, 1).Range.Text = "Vice President Membership";
          officersTable.Cell(11, 2).Range.Text = "Vice President Public Relations";
          officersTable.Rows[5].Range.Font.Bold = 1;
          officersTable.Rows[11].Range.Font.Bold = 1;
          officersTable.Rows[5].Range.Font.Size = 11;
          officersTable.Rows[11].Range.Font.Size = 11;
          officersTable.Rows[4].Range.Font.Size = 8;
          officersTable.Rows[6].Range.Font.Size = 9;
          officersTable.Rows[7].Range.Font.Size = 9;
          officersTable.Rows[8].Range.Font.Size = 9;
          officersTable.Rows[9].Range.Font.Size = 9;
          officersTable.Rows[10].Range.Font.Size = 9;
          officersTable.Rows[12].Range.Font.Size = 9;
          officersTable.Rows[13].Range.Font.Size = 9;
          officersTable.Rows[14].Range.Font.Size = 9;
        }
        // get club officers for that club

        DataSet dsOfficers = new DataSet();

        SqlDataAdapter daOfficers = new SqlDataAdapter("select office, memberid, clubno from clubofficers where clubno = " + clubNo +
          "and office not in ('CTREAS','CSAA','CSEC')", conn);

        daOfficers.Fill(dsOfficers);

        DataTable officersDataTable = dsOfficers.Tables["Table"];

        foreach (DataRow rowOfficer in officersDataTable.Rows)
        {
          bool bFullMemberInfo = false;

          office = rowOfficer["Office"].ToString().Trim();
          memberID = (int)rowOfficer[1];
          email = "";
          name = "";
          loc1 = "";
          loc2 = "";
          phone1 = "";
          phone2 = "";

          //if (office == "CPRES" || office == "CVPE")
          //  bFullMemberInfo = true;

          GenerateMemberInfo(memberID, ref name, ref loc1, ref loc2, ref phone1, ref phone2, ref email, bPrisonClub, bFullMemberInfo, true);
          /*
                    //if (clubNo == "722657")
                    //  name = "";
                    if (clubNo == "7187" || clubNo == "8486" || clubNo == "771553" || clubNo == "814824" || clubNo == "782516" ||
                       clubNo == "722657" || clubNo == "1071907" || clubNo == "1167482")
                      bPrisonClub = true;

                    if (clubNo == "782516" || clubNo == "1071907")
                    {
                      loc1 = "%Shirley Foster, PO Box 2289";
                      loc2 = "Blythe, CA 92226-2349";
                      phone1 = "";
                      phone2 = "";
                      email = "";
                    }

                    //if (clubNo == 
                    //|| clubNo == "1167482"
          
                    if (clubNo == "7184")
                    {
                      loc1 = "16756 Chino Corona Rd";
                      if (office == "CVPM" || office == "CVPPR")
                        phone1 = "";
                    }

                    if (clubNo == "782516")
                      loc1 = "%Shirley Foster, PO Box 2289";
                    */
          if (bPrisonClub)
          {
            if (division == "D")
            {

              if (office == "CPRES")
                officersTable.Cell(7, 1).Range.Text = "Confidential";
              else if (office == "CVPE")
                officersTable.Cell(7, 2).Range.Text = "Confidential";
              else if (office == "CVPM")
                officersTable.Cell(9, 1).Range.Text = "Confidential";
              else if (office == "CVPPR")
                officersTable.Cell(9, 2).Range.Text = "Confidential";
            }
            else
            {
              if (office == "CPRES")
                officersTable.Cell(7, 1).Range.Text = name;
              else if (office == "CVPE")
                officersTable.Cell(7, 2).Range.Text = name;
              else if (office == "CVPM")
                officersTable.Cell(9, 1).Range.Text = name;
              else if (office == "CVPPR")
                officersTable.Cell(9, 2).Range.Text = name;
            }
          }
          else
          {
            if (office == "CPRES")
            {
              officersTable.Cell(6, 1).Range.Text = name.ToString();
              officersTable.Cell(7, 1).Range.Text = loc1;
              officersTable.Cell(8, 1).Range.Text = loc2;
              officersTable.Cell(9, 1).Range.Text = phone1;
              officersTable.Cell(10, 1).Range.Text = email;
            }
            else if (office == "CVPE")
            {
              officersTable.Cell(6, 2).Range.Text = name;
              officersTable.Cell(7, 2).Range.Text = loc1;
              officersTable.Cell(8, 2).Range.Text = loc2;
              officersTable.Cell(9, 2).Range.Text = phone1;
              officersTable.Cell(10, 2).Range.Text = email;
            }
            else if (office == "CVPM")
            {
              officersTable.Cell(12, 1).Range.Text = name;
              officersTable.Cell(13, 1).Range.Text = phone1;
              officersTable.Cell(14, 1).Range.Text = email;
            }
            else if (office == "CVPPR")
            {
              officersTable.Cell(12, 2).Range.Text = name;
              officersTable.Cell(13, 2).Range.Text = phone1;
              officersTable.Cell(14, 2).Range.Text = email;
            }
          }
        }

        clubLocationTable.Select();
        clubLocationTable.Rows.AllowBreakAcrossPages = 0;
        oDoc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = -1;

        Word.Paragraph oPara4;
        oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
        oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);
        oPara4.Range.Font.Size = 8;
        oPara4.Format.SpaceAfter = 0;
      }
    }

    private void clubListingToolStripMenuItem_Click(object sender, EventArgs e)
    {
      oWord = new Word.Application();
      SetUpDocument();
      Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

      DataSet dsClubs = new DataSet();
      SqlDataAdapter daClubs = new SqlDataAdapter("SELECT * FROM Clubs Where Area > 0" +
           " Order by clubno", conn);

      daClubs.Fill(dsClubs);
      DataTable dtClubs = dsClubs.Tables["Table"];

      Word.Table clubListingTitle;
      wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      clubListingTitle = oDoc.Tables.Add(wrdRng, 1, 1, ref oMissing, ref oMissing);
      clubListingTitle.Range.ParagraphFormat.SpaceAfter = 0;
      clubListingTitle.Cell(1, 1).Range.Text = "Club Listing";
      clubListingTitle.Rows[1].Range.Font.Bold = 1;
      clubListingTitle.Rows[1].Range.Font.Size = 14;
      clubListingTitle.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
      //clubListingTitle.Select();
      //oDoc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = 0;

      double clubCount = dtClubs.Rows.Count;
      double rowCount = clubCount / 3.0;

      Word.Table clubListingTable;
      wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      clubListingTable = oDoc.Tables.Add(wrdRng, (int)rowCount + 1, 6, ref oMissing, ref oMissing);
      clubListingTable.Range.ParagraphFormat.SpaceAfter = 0;
      //clubLocationTable.AllowAutoFit = true;
      clubListingTable.BottomPadding = 0;

      int counter = 2;
      int column = 1;
      foreach (DataRow clubListingRow in dtClubs.Rows)
      {
        string clubNo = clubListingRow["clubNo"].ToString().Trim();
        string division = clubListingRow["division"].ToString().Trim();
        string area = clubListingRow["area"].ToString().Trim();
        if (counter > (rowCount + 2))
        {
          column += 2;
          counter = 2;
        }

        clubListingTable.Cell(counter, column).Range.Text = clubNo;
        clubListingTable.Cell(counter, column + 1).Range.Text = division + "" + area;
        counter++;
      }
    }

    private void membersToolStripMenuItem3_Click(object sender, EventArgs e)
    {
      oWord = new Word.Application();
      SetUpDocument();
      Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

      DataSet dsMembers = new DataSet();
      SqlDataAdapter daMembers = new SqlDataAdapter("SELECT * FROM Members " +
           " Order by zip", conn);

      daMembers.Fill(dsMembers);
      DataTable dtMembers = dsMembers.Tables["Table"];

      int memberCount = dtMembers.Rows.Count;

      Word.Table memberLabelTable;
      wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      memberLabelTable = oDoc.Tables.Add(wrdRng, memberCount * 4, 1, ref oMissing, ref oMissing);
      memberLabelTable.Range.ParagraphFormat.SpaceAfter = 0;
      memberLabelTable.BottomPadding = 0;

      string firstName;
      string middleName;
      string lastName;
      string title;
      string mailstop;
      string address1;
      string address2;
      string city;
      string state;
      string zip;
      string name;
      string loc1;
      string loc2;

      int counter = 0;
      foreach (DataRow memberListingRow in dtMembers.Rows)
      {
        firstName = memberListingRow["FirstName"].ToString().Trim();
        middleName = memberListingRow["MiddleName"].ToString().Trim();
        lastName = memberListingRow["LastName"].ToString().Trim();
        title = memberListingRow["Title"].ToString().Trim();
        mailstop = memberListingRow["mailstop"].ToString().Trim();
        address1 = memberListingRow["Address1"].ToString().Trim();
        address2 = memberListingRow["Address2"].ToString().Trim();
        city = memberListingRow["City"].ToString().Trim();
        state = memberListingRow["State"].ToString().Trim();
        zip = memberListingRow["Zip"].ToString().Trim();

        name = firstName;
        if (middleName.Length > 1)
          name += " " + middleName;

        name += " " + lastName;

        if (title.Length > 1)
          name += ", " + title;

        loc1 = "";
        if (mailstop.Length > 1)
          loc1 = mailstop + ", ";

        if (address1.Length > 1)
          loc1 += address1 + " ";

        if (address2.Length > 1)
          loc1 += address2;

        loc2 = "";

        if (city.Length > 1)
          loc2 = city + ", ";

        if (state.Length > 1)
          loc2 += state + " ";

        if (zip.Length > 1)
          loc2 += zip;

        memberLabelTable.Cell(1 + (4 * counter), 1).Range.Text = name;
        memberLabelTable.Cell(2 + (4 * counter), 1).Range.Text = loc1;
        memberLabelTable.Cell(3 + (4 * counter), 1).Range.Text = loc2;
        memberLabelTable.Cell(4 + (4 * counter), 1).Range.Text = "";
        memberLabelTable.Rows[1 + (4 * counter)].Range.Font.Size = 9;
        memberLabelTable.Rows[2 + (4 * counter)].Range.Font.Size = 9;
        memberLabelTable.Rows[3 + (4 * counter)].Range.Font.Size = 9;
        memberLabelTable.Rows[4 + (4 * counter)].Range.Font.Size = 9;

        counter++;
      }
    }

    private void saveToolStripMenuItem_Click(object sender, EventArgs e)
    {
      // Create a new file in C:\\ dir
      XmlTextWriter textWriter = new XmlTextWriter("C:\\myXmlFile.xml", null);

      // Opens the document
      textWriter.WriteStartDocument();

      // Write comments
      textWriter.WriteComment("First Comment XmlTextWriter Sample Example");
      textWriter.WriteComment("myXmlFile.xml in root dir");

      // Write first element
      textWriter.WriteStartElement("Student");
      textWriter.WriteStartElement("r", "RECORD", "urn:record");

      // Write next element
      textWriter.WriteStartElement("Name", "");
      textWriter.WriteString("Student");
      textWriter.WriteEndElement();

      // Write one more element
      textWriter.WriteStartElement("Address", ""); textWriter.WriteString("Colony");
      textWriter.WriteEndElement();

      // WriteChars
      char[] ch = new char[3];
      ch[0] = 'a';
      ch[1] = 'r';
      ch[2] = 'c';
      textWriter.WriteStartElement("Char");
      textWriter.WriteChars(ch, 0, ch.Length);
      textWriter.WriteEndElement();

      // Ends the document.
      textWriter.WriteEndDocument();

      // close writer
      textWriter.Close();
    }

    private void OnNodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
    {
      if (e.Node.Parent == null)
        return;

      switch (e.Node.Parent.Text)
      {
        case "Chairs":
          {
            DataSet dsChairs = new DataSet();
            SqlDataAdapter daChairs = new SqlDataAdapter("SELECT Chairs.Chair, Members.FirstName, Members.MiddleName,  Members.LastName, " +
                "Members.Title, Members.Mailstop, Members.Address1,Members.Address2, Members.City, Members.Zip, Members.WorkPhone, Members.HomePhone, " +
                "Members.CellPhone, Members.Email " +
                " FROM    Chairs INNER JOIN  Members ON Chairs.MemberID = Members.MemberID", conn);
            daChairs.Fill(dsChairs);
            DataTable dtChairs = dsChairs.Tables["Table"];

            string office;
            string firstName;
            string middleName;
            string lastName;
            string title;
            string mailstop;
            string address1;
            string address2;
            string city;
            string zip;
            string wkphone;
            string hmphone;
            string cellphone;
            string email;
            string name;
            string loc1;
            string loc2;
            string phone1;
            string phone2;

            foreach (DataRow rowChairs in dtChairs.Rows)
            {
              office = rowChairs.ItemArray[0].ToString().Trim();
              firstName = rowChairs.ItemArray[1].ToString().Trim();
              middleName = rowChairs.ItemArray[2].ToString().Trim();
              lastName = rowChairs.ItemArray[3].ToString().Trim();
              title = rowChairs.ItemArray[4].ToString().Trim();
              //mailstop = rowChairs.ItemArray[5].ToString().Trim();
              address1 = rowChairs.ItemArray[6].ToString().Trim();
              address2 = rowChairs.ItemArray[7].ToString().Trim();
              city = rowChairs.ItemArray[8].ToString().Trim();
              zip = rowChairs.ItemArray[9].ToString().Trim();
              wkphone = rowChairs.ItemArray[10].ToString().Trim();
              hmphone = rowChairs.ItemArray[11].ToString().Trim();
              cellphone = rowChairs.ItemArray[12].ToString().Trim();
              email = rowChairs.ItemArray[13].ToString().Trim();
            }
          }
          break;
        case "Officers":
          break;

      }
    }

    private void listViewOfficer_SelectedIndexChanged(object sender, EventArgs e)
    {

    }

    private void membersupdateToolStripMenuItem_Click(object sender, EventArgs e)
    {
      string CreateMembersTableSQL = "CREATE TABLE Members_Jan08 (MemberID INT PRIMARY KEY NOT NULL,"
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
      + "WorkPhone varchar(25),"
      + "HomePhone varchar(20),"
      + "CellPhone varchar(20),"
      + "Email varchar(40),"
      + "Email2 varchar(40),"
      + "Web varchar(50))";

      SqlCommand DBCmd = new SqlCommand(CreateMembersTableSQL, conn);
      //DBCmd.CommandText = CreateMembersTableSQL;
      DBCmd.ExecuteNonQuery();

    }

    private void clubMembersupdateToolStripMenuItem_Click(object sender, EventArgs e)
    {
      string CreateClubMembersTable = "CREATE TABLE ClubMembers_Jan08 (ClubMemberID INT IDENTITY(1,1) PRIMARY KEY NOT NULL,"
      + "ClubNo int,"
      + "MemberID int)";

      SqlCommand dbCmd = new SqlCommand(CreateClubMembersTable, conn);
      dbCmd.ExecuteNonQuery();

    }

    private void clubOfficersupdateToolStripMenuItem_Click(object sender, EventArgs e)
    {
      string CreateClubOfficersTableSQL = "CREATE TABLE ClubOfficers_Update (ClubOfficerID INT IDENTITY(1,1) PRIMARY KEY NOT NULL,"
      + "ClubNo  int,"
      + "Office varchar(10),"
      + "MemberID int)";

      SqlCommand DBCmd = new SqlCommand(CreateClubOfficersTableSQL, conn);
      //DBCmd.CommandText = CreateClubOfficersTableSQL;
      DBCmd.ExecuteNonQuery();
    }

    private void othersupdateToolStripMenuItem_Click(object sender, EventArgs e)
    {
      string CreateDistrictOfficersTable = "CREATE TABLE DistrictOfficers_Update (DistrictOfficers  smallint IDENTITY(1,1) PRIMARY KEY NOT NULL,"
      + "memberID int NOT NULL,"
      + "Office varchar(50) NOT NULL)";

      SqlCommand dbCmd = new SqlCommand(CreateDistrictOfficersTable, conn);
      dbCmd.ExecuteNonQuery();

      //string CreateClubTable = "CREATE TABLE Clubs_Update (ClubID INT IDENTITY(1,1) NOT NULL,"
      //+ "ClubNo int PRIMARY KEY NOT NULL,"
      //+ "ClubName  varchar(max),"
      //+ "Days varchar(50),"
      //+ "Time char(10),"
      //+ "Website varchar(max),"
      //+ "Phone varchar(25),"
      //+ "Email varchar(50),"
      //+ "Location1 varchar(50),"
      //+ "Address varchar(50),"
      //+ "City varchar(50),"
      //+ "Zip char(10),"
      //+ "Area char(2),"
      //+ "Division char(1))";

      //dbCmd.CommandText = CreateClubTable;
      //dbCmd.ExecuteNonQuery();
    }

    private void clubupdateToolStripMenuItem_Click(object sender, EventArgs e)
    {
      DistrictManagerEngine.DMLoader DML = new DistrictManagerEngine.DMLoader();
      DML.LoadClubs();

      //FileStream fleReader = new FileStream("E:\\TI\\Directory\\Nov_07\\club_update_nov_07.csv", FileMode.Open, FileAccess.Read);
      //StreamReader stmReader = new StreamReader(fleReader);

      //string line;// = stmReader.ReadLine(); //skip header

      //char[] delims = new char[] { ',' };
      //while ((line = stmReader.ReadLine()) != null)
      //{
      //  string[] pole = line.Split(delims, StringSplitOptions.None);
      //  //clubRecord rcd = new clubRecord(pole);

      //  //string club = "INSERT INTO Clubs_Update VALUES (" + rcd.ClubNo + ",'" + rcd.ClubName + "','" + rcd.Days + "','" + rcd.Time + "','" + rcd.Website + "' ,'" +
      //  //            rcd.Phone + "','" + rcd.Email + "','" + rcd.Location1 + "','" + rcd.Location2 + "','" + rcd.Address + "','" + rcd.City + "','" + rcd.Zip + "','" + rcd.Area + "','" + rcd.Division + "')";
      //  //  Insert dept table records first
      //  //SqlCommand dbcmd = new SqlCommand(club, conn);
      //  //dbcmd.ExecuteNonQuery();
      //}
    }

    private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
    {

    }


    private void Form1_FormClosing(object sender, FormClosingEventArgs e)
    {
      // save position and size of controls for start up
      string fileName = Application.LocalUserAppDataPath + @"\MainForm.txt";
      using
      (
        StreamWriter writer = new StreamWriter(fileName))
      {
        FormWindowState state = this.WindowState;
        writer.WriteLine(this.Location.ToString());
        writer.WriteLine(this.ClientSize.ToString());
      }
    }

    private void treeView1_DragEnter(object sender, DragEventArgs e)
    {
      e.Effect = DragDropEffects.Move;

      //string stuff = e.Data.ToString();
    }

    private void treeView1_DragDrop(object sender, DragEventArgs e)
    {

    }

    private void treeView1_DragOver(object sender, DragEventArgs e)
    {
      e.Effect = DragDropEffects.Move;

    }

    private void treeView1_MouseDown(object sender, MouseEventArgs e)
    {
      DoDragDrop(sender.ToString(), DragDropEffects.Copy);
    }

    void SetDropEffect(DragEventArgs e)
    {
      if (e.Data.GetDataPresent(typeof(string)))
      {
        e.Effect = DragDropEffects.Move;
      }
    }

    public void GenerateMemberInfo(int memberID, ref string name, ref string location, ref string location2,
                            ref string phone1, ref string phone2, ref string email, bool prisonClub, bool full, bool clubOfficer)
    {
      string firstName = "";
      string middleName = "";
      string lastName = "";
      string title = "";
      string address1 = "";
      string address2 = "";
      string address3 = "";
      string city = "";
      string zip = "";
      string wkphone = "";
      string hmphone = "";
      string cellphone = "";
      string loc1 = "";
      string loc2 = "";

      DataSet dsMember = new DataSet();
      SqlDataAdapter daMember = new SqlDataAdapter();
      SqlCommand clubOfficerCmd = new SqlCommand("select firstname, middlename, lastname, title, " +
        "workphone, homephone, cellphone, email from members " +
        "where memberid = " + memberID, conn);

      SqlCommand otherOfficerCmd = new SqlCommand("select firstname, middlename, lastname, title, " +
        "workphone, homephone, cellphone, email from members " +
        "where memberid = " + memberID, conn);

      if (clubOfficer)
        daMember.SelectCommand = clubOfficerCmd;
      else
        daMember.SelectCommand = otherOfficerCmd;

      daMember.Fill(dsMember);
      DataTable dtMember = dsMember.Tables[0];
      if (dtMember.Rows.Count < 1)
        return;
      DataRow rowMember = dtMember.Rows[0];

      if (rowMember != null)
      {
        firstName = rowMember.ItemArray[0].ToString().Trim();
        middleName = rowMember.ItemArray[1].ToString().Trim();
        lastName = rowMember.ItemArray[2].ToString().Trim();
        title = rowMember.ItemArray[3].ToString().Trim();
        //address1 = rowMember.ItemArray[4].ToString().Trim();
        //address2 = rowMember.ItemArray[5].ToString().Trim();
        //address3 = rowMember.ItemArray[6].ToString().Trim();
        //city = rowMember.ItemArray[7].ToString().Trim();
        //zip = rowMember.ItemArray[8].ToString().Trim();
        wkphone = rowMember.ItemArray[4].ToString().Trim();
        hmphone = rowMember.ItemArray[5].ToString().Trim();
        cellphone = rowMember.ItemArray[6].ToString().Trim();
        email = rowMember.ItemArray[7].ToString().Trim();

        name = firstName;
        if (prisonClub)
        {
          String temp = lastName.Remove(1);
          lastName = temp;
          lastName += ".";
        }
        name += " " + lastName;

        if (title.Length > 1)
          name += ", " + title;

        //if (address1.Length > 1)
        //  location = address1;

        //if (address2.Length > 0)
        //  location += ", " + address2;

        //if (address3.Length > 0)
        //  location += ", " + address3;

        //if (city.Length > 1)
        //  location2 += city + ", ";

        //location2 += "CA ";

        //if (zip.Length > 1)
        //  location2 += zip;

        ModifyPhoneNumbers(ref wkphone, ref hmphone, ref cellphone);

        wkphone = "";

        if (hmphone.Length > 1 && cellphone.Length > 1)
          phone1 = "H " + hmphone + ", C " + cellphone;
        else if (hmphone.Length < 1 && cellphone.Length > 1)
          phone1 = "C " + cellphone;
        else if (hmphone.Length > 1 && cellphone.Length < 1)
          phone1 = "H " + hmphone;
        /*
        if (wkphone.Length > 1 && hmphone.Length > 1 && cellphone.Length > 1)
        {
          phone1 = "W " + wkphone + ", H " + hmphone;
          phone2 = "C " + cellphone;
        }
        else if (wkphone.Length > 1 && hmphone.Length < 1 && cellphone.Length > 1)
        {
          phone1 = "W " + wkphone + ", C " + cellphone;
          if (full && clubOfficer)
          {
            phone2 = email;
            email = "";
          }
          else
            phone2 = "";
        }
        else if (wkphone.Length < 1 && hmphone.Length > 1 && cellphone.Length > 1)
        {
          phone1 = "H " + hmphone + ", C " + cellphone;
          if (full && clubOfficer)
          {
            phone2 = email;
            email = "";
          }
          else
            phone2 = "";
        }
        else if (wkphone.Length < 1 && hmphone.Length < 1 && cellphone.Length > 1)
        {
          phone1 = "C " + cellphone;
          if (full && clubOfficer)
          {
            phone2 = email;
            email = "";
          }
          else
            phone2 = "";
        }
        else if (wkphone.Length > 1 && hmphone.Length > 1 && cellphone.Length < 1)
        {
          phone1 = "W " + wkphone + ", H " + hmphone;
          if (full && clubOfficer)
          {
            phone2 = email;
            email = "";
          }
          else
            phone2 = "";
        }
        else if (wkphone.Length < 1 && hmphone.Length > 1 && cellphone.Length < 1)
        {
          phone1 = "H " + hmphone;
          if (full && clubOfficer)
          {
            phone2 = email;
            email = "";
          }
          else
            phone2 = "";
        }
        else if (wkphone.Length > 1 && hmphone.Length < 1 && cellphone.Length < 1)
        {
          phone1 = "W " + wkphone;
          if (full && clubOfficer)
          {
            phone2 = email;
            email = "";
          }
          else
            phone2 = "";
        }
        */
        //if (division == "C" && area == 6)
        //{
        //  name = firstName;
        //  if (title.Length > 1)
        //    name += ", " + title;

        //  loc1 = "% Ms S Hoover - Felon Records";
        //  phone1 = loc2;
        //  loc2 = "PO Box 1841";
        //  phone1 = "dzeller@ca.rr.com";
        //  phone2 = "";
        //  email = "";
        //}
      }
    }

    private void ModifyPhoneNumbers(ref String worknum, ref String homenum, ref String cellnum)
    {
      if (worknum.Length < 10)
        worknum = "";
      else if (worknum.IndexOf('.') < 0 && worknum.Length > 9)
      {
        String areaCode = worknum.Substring(0, 3);
        String exchange = worknum.Substring(3, 3);
        String num = worknum.Substring(6, 4);
        String newNum = areaCode + "." + exchange + "." + num;

        String extension;
        if (worknum.Length > 10) // has extension
        {
          extension = worknum.Substring(10, (worknum.Length - 10));
          if (extension.StartsWith("x") || extension.StartsWith("ext") || extension.StartsWith("ext."))
          {
            newNum += " " + extension;
          }
          else newNum += " x" + extension;
        }
        worknum = newNum;

      }

      if (homenum.Length < 10)
        homenum = "";
      else if (homenum.IndexOf('.') < 0 && homenum.Length > 9)
      {
        String areaCode = homenum.Substring(0, 3);
        String exchange = homenum.Substring(3, 3);
        String num = homenum.Substring(6, 4);
        String newNum = areaCode + "." + exchange + "." + num;

        String extension;
        if (homenum.Length > 10) // has extension
        {
          extension = homenum.Substring(10, (homenum.Length - 10));
          newNum += " " + extension;
        }
        homenum = newNum;
      }

      if (cellnum.Length < 10)
        cellnum = "";
      else if (cellnum.IndexOf('.') < 0 && cellnum.Length > 9)
      {
        String areaCode = cellnum.Substring(0, 3);
        String exchange = cellnum.Substring(3, 3);
        String num = cellnum.Substring(6, 4);
        String newNum = areaCode + "." + exchange + "." + num;

        String extension;
        if (cellnum.Length > 10) // has extension
        {
          extension = cellnum.Substring(10, (cellnum.Length - 10));
          newNum += " " + extension;
        }
        cellnum = newNum;

      }


    }

    private void toolStripMenuItem2_Click(object sender, EventArgs e)
    {
      division = "E";
      Thread divThread = new Thread(new ThreadStart(GenerateDivisionThreadStart));
      divThread.Start();
    }

    private void toolStripMenuItem4_Click(object sender, EventArgs e)
    {
      oWord = new Word.Application();
      SetUpDocument();
      //GenerateArea("D", 5);
      GenerateAreaNew(oDoc, "D", 5);
    }

    private void byAreasToolStripMenuItem_Click(object sender, EventArgs e)
    {
      DataSet dsDivisions = new DataSet();
      SqlDataAdapter daDiv = new SqlDataAdapter("Select * from DivAreaMatrix", conn);

      daDiv.Fill(dsDivisions);
      DataTable dtDivisions = dsDivisions.Tables["Table"];
      foreach (DataRow rowDivision in dtDivisions.Rows)
      {
        //for (int area = 1; area <= (int)rowDivision.ItemArray[1]; area++)
        //{
        oWord = new Word.Application();
        SetUpDocument();
        GenerateDivision(rowDivision.ItemArray[1].ToString());
        //GenerateArea(rowDivision.ItemArray[0], area);

        //Word.Table tTable;

        //int nTables = oDoc.Tables.Count;
        //if (nTables < 1)
        //  return;

        //for (int index = 1; index <= nTables; index++)
        //{
        //  tTable = oDoc.Tables[index];
        //  tTable.Select();
        //  tTable.Rows.AllowBreakAcrossPages = 0;
        //  oDoc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = -1;

        //  //for (int rowIndex = 1; rowIndex < tTable.Rows.Count; rowIndex++)
        //  //  tTable.Rows[rowIndex].Range.ParagraphFormat.KeepWithNext = -1;
        //}
        //}
      }
    }

    private void newFormatToolStripMenuItem_Click(object sender, EventArgs e)
    {
      oWord = new Word.Application();
      //SetUpDocument();
      DataSet dsDivisions = new DataSet();
      SqlDataAdapter daDiv = new SqlDataAdapter("Select * from DivAreaMatrix", conn);

      daDiv.Fill(dsDivisions);
      DataTable dtDivisions = dsDivisions.Tables["Table"];
      foreach (DataRow rowDivision in dtDivisions.Rows)
      {
        //for (int area = 1; area <= (int)rowDivision.ItemArray[1]; area++)
        //{
        oWord = new Word.Application();
        SetUpDocument();
        GenerateDivision(rowDivision.ItemArray[1].ToString());
      }
      //GenerateAreaNew("D", 1);
    }

    private void GenerateAreaNew(Word._Document doc, string division, int area)
    {
      Word.Range wrdRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;

      DataSet dsAreaGov = new DataSet();
      SqlDataAdapter daAreaGov = new SqlDataAdapter("select Area, MemberID, Status, Email from AreaGovernor " +
        "where Division = " + "'" + division + "'" + "AND AreaGovernor.Area = " + area, conn);
      int memberID = 0;

      daAreaGov.Fill(dsAreaGov);
      DataTable datatableAreaGov = dsAreaGov.Tables["Table"];
      DataRow rowAreaGov = null;
      String areaStatus = "";
      String email = "";

      if (datatableAreaGov.Rows.Count > 0)
      {
        rowAreaGov = datatableAreaGov.Rows[0];
        memberID = (int)rowAreaGov.ItemArray[1];
        if (! System.DBNull.Value.Equals(rowAreaGov.ItemArray[2]))
          areaStatus = (string)rowAreaGov.ItemArray[2];
        if (!System.DBNull.Value.Equals(rowAreaGov.ItemArray[3]))
          email = (string)rowAreaGov.ItemArray[3];
        else
          email = division + area.ToString() + "Dir@d12toastmasters.org";
      }

      string office = "";

      //string state = "CA";
      string name = "";
      string loc1 = "";
      string loc2 = "";
      string phone1 = "";
      string phone2 = "";
      string blank = "";

      if (memberID > 0)
        GenerateMemberInfo(memberID, ref name, ref loc1, ref loc2, ref phone1, ref phone2, ref blank, false, true, false);

      Word.Table areaGovTitle;
      wrdRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;

      bool distinguishedArea = false;
      int offset = 0;
      if (areaStatus.Contains("S") || areaStatus.Contains("P") || areaStatus.Contains("D"))
      {
        areaGovTitle = oDoc.Tables.Add(wrdRng, 2, 1, ref oMissing, ref oMissing);
        offset = 1;
        distinguishedArea = true;
      }
      else
        areaGovTitle = oDoc.Tables.Add(wrdRng, 1, 1, ref oMissing, ref oMissing);

      areaGovTitle.Range.ParagraphFormat.SpaceAfter = 0;
      areaGovTitle.Cell(1, 1).Range.Text = "Area " + division + area;
      areaGovTitle.Rows[1].Range.Font.Bold = 1;
      areaGovTitle.Rows[1].Range.Font.Size = 14;
      // areaGovTitle.Rows[2].Range.Font.Size = 8;
      areaGovTitle.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
      //areaGovTitle.Select();
      //oDoc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = 0;

      if (distinguishedArea)
      {
        string areaStat = "2014-2015 ";
        switch (areaStatus)
        {
          case "P":
            areaStat += "President's" + " ";
            break;
          case "S":
            areaStat += "Select" + " ";
            break;
        }
        areaStat += "Distinguished Area";

        areaGovTitle.Cell(2, 1).Range.Text = areaStat;

        areaGovTitle.Rows[2].Range.Font.Bold = 1;
        areaGovTitle.Rows[2].Range.Font.Size = 9;
      }

      DataSet dsAssistantAreaGov = new DataSet();
      SqlDataAdapter daAssistantAreaGov = new SqlDataAdapter("select Area, MemberID from AssistantAreaGovernor " +
        "where Division = " + "'" + division + "'" + "AND Area = " + area + " and memberID > 0", conn);

      daAssistantAreaGov.Fill(dsAssistantAreaGov);

      DataTable datatableAssistantAreaGov = dsAssistantAreaGov.Tables["Table"];

      int AreaAssistantCount = datatableAssistantAreaGov.Rows.Count;

      // populate Area governor and assistant info
      Word.Table areaGovernorTable;
      wrdRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      if (AreaAssistantCount < 2)
        areaGovernorTable = doc.Tables.Add(wrdRng, 4, 2, ref oMissing, ref oMissing);
      else
      {
        areaGovernorTable = doc.Tables.Add(wrdRng, 9, 2, ref oMissing, ref oMissing);
        areaGovernorTable.Rows[6].Range.Font.Size = 7;
      }

      areaGovernorTable.Range.ParagraphFormat.SpaceAfter = 0;
      //clubLocationTable.AllowAutoFit = true;
      areaGovernorTable.BottomPadding = 0;
      //clubLocationTable.Rows[1].Height = 15;

      //if (phone2.Length < 1 && email.Length > 1)
      //{
      //  phone2 = email;
      //  email = "";
      //}

      areaGovernorTable.Cell(2 + offset, 1).Range.Text = "Area Director";
      areaGovernorTable.Cell(2 + offset, 2).Range.Text = "Assistant Area Director";
      areaGovernorTable.Rows[2 + offset].Range.Font.Bold = 1;
      areaGovernorTable.Rows[2 + offset].Range.Font.Size = 11;
      areaGovernorTable.Rows[3 + offset].Range.Font.Size = 9;
      areaGovernorTable.Rows[4 + offset].Range.Font.Size = 9;
      areaGovernorTable.Rows[5 + offset].Range.Font.Size = 9;
      areaGovernorTable.Cell(3 + offset, 1).Range.Text = name;
      //areaGovernorTable.Cell(4, 1).Range.Text = loc1;
      //areaGovernorTable.Cell(5, 1).Range.Text = loc2;
      areaGovernorTable.Cell(4 + offset, 1).Range.Text = phone1;
      //areaGovernorTable.Cell(5 + offset, 1).Range.Text = phone2;
      areaGovernorTable.Cell(5 + offset, 1).Range.Text = email;

      //DataRow rowAsstAreaGov = null;
      memberID = 0;
      int interRow = 0;
      int interCol = 2;
      foreach (DataRow rowAsstAreaGov in datatableAssistantAreaGov.Rows)
      {
        memberID = (int)rowAsstAreaGov.ItemArray[1];
        //if (datatableAssistantAreaGov.Rows.Count > 0)
        //{
        //  rowAsstAreaGov = datatableAssistantAreaGov.Rows[0];
        //  memberID = (int)rowAsstAreaGov.ItemArray[1];
        //}
        // empty strings

        email = "";
        name = "";
        loc1 = "";
        loc2 = "";
        phone1 = "";
        phone2 = "";

        if (memberID > 0)
          GenerateMemberInfo(memberID, ref name, ref loc1, ref loc2, ref phone1, ref phone2, ref email, false, true, false);

        //if (phone2.Length < 1 && email.Length > 1)
        //{
        //  phone2 = email;
        //  email = "";
        //}

        if (interRow > 0)
        {
          string bloodyStemplerText = "Assistant Area Governor";

          areaGovernorTable.Cell(3 + (3 * interRow) + offset + interRow, 1).Range.Text = bloodyStemplerText;
          areaGovernorTable.Rows[3 + (3 * interRow) + offset + interRow].Range.Font.Bold = 1;
          areaGovernorTable.Rows[3 + (3 * interRow) + offset + interRow].Range.Font.Size = 11;
          if (AreaAssistantCount == 3)
          {
            areaGovernorTable.Cell(3 + (3 * interRow) + offset + interRow, 2).Range.Text = "Assistant Area Governor";
          }
        }

        areaGovernorTable.Cell((3 + (3 * interRow) + offset + (interRow * 2)), interCol).Range.Text = name;
        //areaGovernorTable.Cell(4, 2).Range.Text = loc1;
        //areaGovernorTable.Cell(5, 2).Range.Text = loc2;
        areaGovernorTable.Cell((4 + (3 * interRow) + offset + (interRow * 2)), interCol).Range.Text = phone1;
        //areaGovernorTable.Cell((5 + (3 * interRow) + offset + (interRow * 2)), interCol).Range.Text = phone2;
        areaGovernorTable.Cell((5 + (3 * interRow) + offset + (interRow * 2)), interCol).Range.Text = email;

        //areaGovernorTable.Rows[(2 + (3 * interRow) + offset + (interRow * 2))].Range.Font.Bold = 1;
        //areaGovernorTable.Rows[(2 + (3 * interRow) + offset + (interRow * 2))].Range.Font.Size = 11;
        areaGovernorTable.Rows[(3 + (3 * interRow) + offset + (interRow * 2))].Range.Font.Size = 9;
        areaGovernorTable.Rows[(4 + (3 * interRow) + offset + (interRow * 2))].Range.Font.Size = 9;
        areaGovernorTable.Rows[(5 + (3 * interRow) + offset + (interRow * 2))].Range.Font.Size = 9;
        //areaGovernorTable.Rows[(6 + (3 * interRow) + offset + (interRow * 2))].Range.Font.Size = 9;
        //areaGovernorTable.Rows[7].Range.Font.Size = 9;
        //areaGovernorTable.Rows[8].Range.Font.Size = 9;
        if (interRow < 1)
          interRow++;

        if (interCol == 2)
          interCol = 1;
        else if (interCol == 1)
          interCol = 2;

      }

      datatableAssistantAreaGov.Clear();

      areaGovernorTable.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
      areaGovernorTable.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
      areaGovernorTable.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
      areaGovernorTable.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
      areaGovernorTable.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth100pt;

      Word.Paragraph areaGovBreakafter;
      object oRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
      areaGovBreakafter = doc.Content.Paragraphs.Add(ref oRng);
      areaGovBreakafter.Range.Font.Size = 4;
      areaGovBreakafter.Format.SpaceAfter = 0;

      //areaGovernorTable.Select();
      //areaGovernorTable.AllowPageBreaks = false;

      // empty strings

      email = "";
      name = "";
      loc1 = "";
      loc2 = "";
      phone1 = "";
      phone2 = "";

      // Clubs in Area
      DataSet dsClubs = new DataSet();
      SqlDataAdapter daClubs = new SqlDataAdapter("SELECT * FROM Clubs " +
          " WHERE Division = " + "'" + division + "'" + " AND Area = " + "'" + area + "'" +
           " Order by clubno", conn);

      daClubs.Fill(dsClubs);
      DataTable dataTableClub = dsClubs.Tables["Table"];
      string clubNo;
      string clubName;
      string dayOfTheWeek;
      string time;
      string web = string.Empty;
      string address;
      //string frequency;
      string phone = string.Empty;
      string city;
      string zip;
      //string email2;
      string clubStatus = string.Empty;
      string meeting = string.Empty;
      string facebook = string.Empty;
      foreach (DataRow rowClub in dataTableClub.Rows)
      {
        //object clubNumber = rowClub.ItemArray[0];

        // populate club info
        // add a table for the club information
        clubNo = rowClub.ItemArray[1].ToString().Trim();
        bool bSEC = false;

        //if (division == "A")
        //{
        //  if (clubNo == "2485" || clubNo == "3014")
        //    bSEC = true;
        //}

        //if (division == "C")
        //{
        //  if (clubNo == "6900" || clubNo == "7730")
        //    bSEC = true;
        //}

        if (division == "D")
        {
          if (clubNo == "782516" || clubNo == "1071907" || clubNo == "9681" ||
              clubNo == "1470790")
            bSEC = true;
        }

        if (division == "E")
        {
          if (clubNo == "7187" || clubNo == "8486" || clubNo == "771553" || clubNo == "814824" ||
              clubNo == "722657" || clubNo == "1535248")
            bSEC = true;
        }


        clubName = rowClub.ItemArray[2].ToString().Trim();
        dayOfTheWeek = rowClub.ItemArray[7].ToString().Trim();
        time = rowClub.ItemArray[8].ToString().Trim();
        if (!System.DBNull.Value.Equals(rowClub.ItemArray[15]))
          web = rowClub.ItemArray[15].ToString().Trim();
        if (!System.DBNull.Value.Equals(rowClub.ItemArray[12]))
          phone = rowClub.ItemArray[12].ToString().Trim();
        //phone2 = rowClub.ItemArray[15].ToString().Trim();
        if (!System.DBNull.Value.Equals(rowClub.ItemArray[13]))
          email = rowClub.ItemArray[13].ToString().Trim();
        //email2 = rowClub.ItemArray[17].ToString().Trim();
        loc1 = rowClub.ItemArray[5].ToString().Trim();
        //loc2 = rowClub.ItemArray[9].ToString().Trim();
        address = rowClub.ItemArray[6].ToString().Trim();
        city = rowClub.ItemArray[9].ToString().Trim();
        zip = rowClub.ItemArray[10].ToString().Trim();
        if (!System.DBNull.Value.Equals(rowClub.ItemArray[14]))
          facebook = rowClub.ItemArray[14].ToString().Trim();
        //frequency = rowClub.ItemArray[7].ToString().Trim();
        clubStatus = rowClub.ItemArray[17].ToString();
        //meeting = rowClub.ItemArray[22].ToString();

        Word.Table clubLocationTable;
        wrdRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
        bool distinguishedClub = false;
        if (clubStatus.Contains("S") || clubStatus.Contains("P") || clubStatus.Contains("D"))
        {
          clubLocationTable = doc.Tables.Add(wrdRng, 4, 1, ref oMissing, ref oMissing);
          distinguishedClub = true;
        }
        else if (meeting == "N")
          clubLocationTable = doc.Tables.Add(wrdRng, 2, 1, ref oMissing, ref oMissing);
        else
          clubLocationTable = doc.Tables.Add(wrdRng, 3, 1, ref oMissing, ref oMissing);
        clubLocationTable.Range.ParagraphFormat.SpaceAfter = 0;
        //clubLocationTable.AllowAutoFit = true;
        //clubLocationTable.BottomPadding = 0;

        // select table to keep from splitting across page breaks
        //int LastRow = clubLocationTable.Range.Rows.Count;
        //Word.Range MyRange = clubLocationTable.Range.Rows[0].Range;
        //MyRange.SetRange(0, LastRow);
        //MyRange = ActiveDocument.Tables(1).Range.Rows(2).Range
        //MyRange.SetRange Start:=MyRange.Start, _
        //End:=ActiveDocument.Tables(1).Range.Rows(LastRow).Range.End 
        //clubLocationTable.Range.SetRange
        //clubLocationTable.Select();
        //oDoc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = -1;

        clubLocationTable.AllowPageBreaks = false;
        //clubLocationTable.Selection.ParagraphFormat.KeepWithNext = true;
        //MyRange.ParagraphFormat.KeepWithNext = true;
        //areaGovTitle.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
        //clubLocationTable.Rows[1].Height = 15;
        string clubRow1;
        string clubRow2 = "";
        string clubRow3 = "";
        string clubRow4 = "";
        string clubRow5 = "";

        // need logic in club info, if something is empty we don't want a bunch of commas
        clubRow1 = "Club " + clubNo + " - " + clubName;
        //clubRow2 = dayOfTheWeek + " " + time + " " + web + " " + phone + " " + email;
        //if (dayOfTheWeek.Length > 1)
        //  clubRow2 += dayOfTheWeek + " ";

        if (dayOfTheWeek.Length > 1)
          clubRow2 += dayOfTheWeek;

        if (time.Length > 1)
          clubRow2 += " " + time;

        if (web.Length > 1)
          clubRow2 += ", " + web;

        if (phone.Length > 1)
          clubRow2 += ", " + phone;

        if (email.Length > 1)
          clubRow2 += ", " + email;

        if (facebook.Length > 1)
          clubRow2 += ", " + facebook;

        //clubRow3 = loc1 + ", " + loc2 + ", " + address + ", " + city + ", " + zip;
        if (loc1.Length > 1)
          clubRow3 += loc1;

        //if (loc2.Length > 1)
        //  clubRow3 += loc2 + ", ";

        if (address.Length > 1)
          clubRow3 += ", " + address;

        if (city.Length > 1)
          clubRow3 += ", " + city;

        if (zip.Length > 1)
          clubRow3 += ", " + zip;

        if (bSEC)
        {
          clubRow2 = "Special Environment Club - contact Area Director";
          clubRow3 = "";
          /*
          if (clubNo == "7187" || clubNo == "8486" || clubNo == "771553" || clubNo == "814824" ||
              clubNo == "722657")
          {
            clubRow2 = "";
            if (dayOfTheWeek.Length > 1)
              clubRow2 += dayOfTheWeek + " ";

            if (time.Length > 1)
              clubRow2 += time + " ";

            clubRow3 = "Contact: Special Environment Clubs Chair, Randy Amelino, DTM";
            clubRow4 = "Email: chemdryall@gnww.net Phone: 951-258-6901";

            if (clubNo == "7187" || clubNo == "8486")
              clubRow2 += "California Institute for Women, 16756 Chino-Corona Rd, Corona, 92880-9508";
            else if (clubNo == "771553" || clubNo == "814824")
              clubRow2 += "California Rehab Center, Bldg 601 Men’s Education Dept, Norco, 92860";
            else if (clubNo == "722657")
              clubRow2 += "California Rehab Center, Bldg 601 Men’s Education Dept, Norco, 92860";
          }

          if (clubNo == "782516" || clubNo == "1071907" || clubNo == "9681" || clubNo == "1167482" ||
              clubNo == "1470790")
          {
            clubRow2 = "";
            if (dayOfTheWeek.Length > 1)
              clubRow2 += dayOfTheWeek + " ";

            if (time.Length > 1)
              clubRow2 += time + " ";

            clubRow3 = "";
            clubRow4 = "";
            clubRow5 = "";

            if (clubNo == "1167482")
              clubRow2 += "Tameka Roberson, 760.922.0680x5074";
            else
              clubRow2 += "Chuckwalla Staff, cindy.nepusz@cdcr.ca.gov, 760.922.5300x5124";


          }
           */
        }

        clubLocationTable.Cell(1, 1).Range.Text = clubRow1;
        clubLocationTable.Rows[1].Range.Font.Bold = 1;
        clubLocationTable.Rows[1].Range.Font.Size = 11;
        //if (meeting == "N")
        //{
        //  clubLocationTable.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
        //  clubLocationTable.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
        //  clubLocationTable.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
        //  clubLocationTable.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
        //  clubLocationTable.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth100pt;
        //  Word.Paragraph locreakafter;
        //  oRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
        //  locreakafter = doc.Content.Paragraphs.Add(ref oRng);
        //  locreakafter.Range.Font.Size = 4;
        //  locreakafter.Format.SpaceAfter = 0;
        //  continue;
        //}
        if (distinguishedClub)
        {
          string clubStat = "2014-2015 ";
          switch (clubStatus)
          {
            case "P":
              clubStat += "President's" + " ";
              break;
            case "S":
              clubStat += "Select" + " ";
              break;
          }
          clubStat += "Distinguished Club";

          clubLocationTable.Cell(2, 1).Range.Text = clubStat;
          clubLocationTable.Cell(3, 1).Range.Text = clubRow2;
          clubLocationTable.Cell(4, 1).Range.Text = clubRow3;

          clubLocationTable.Rows[2].Range.Font.Bold = 1;
          clubLocationTable.Rows[2].Range.Font.Size = 9;
          clubLocationTable.Rows[3].Range.Font.Size = 9;
          clubLocationTable.Rows[4].Range.Font.Size = 9;
        }
        else
        {
          clubLocationTable.Cell(2, 1).Range.Text = clubRow2;
          clubLocationTable.Cell(3, 1).Range.Text = clubRow3;
          clubLocationTable.Rows[2].Range.Font.Size = 9;
          clubLocationTable.Rows[3].Range.Font.Size = 9;
        }

        Word.Paragraph clubLocBreakafter;
        oRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
        clubLocBreakafter = doc.Content.Paragraphs.Add(ref oRng);
        clubLocBreakafter.Range.Font.Size = 1;
        clubLocBreakafter.Format.SpaceAfter = 0;
        clubLocBreakafter.KeepWithNext = -1;

        // determine which officer is selected
        Word.Table officersTable;

        wrdRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
        officersTable = doc.Tables.Add(wrdRng, 7, 2, ref oMissing, ref oMissing);
        officersTable.Range.ParagraphFormat.SpaceAfter = 0;
        officersTable.Columns[1].PreferredWidth = 35F;
        officersTable.Columns[2].PreferredWidth = 290F;
        officersTable.Rows.HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAuto;
        officersTable.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
        officersTable.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
        officersTable.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
        officersTable.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
        officersTable.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth100pt;

        //officersTable.Cell(4, 1).Range.Font.Size = 4;
        //officersTable.Cell(5, 1).Range.Font.Size = 9;
        //officersTable.Cell(6, 1).Range.Font.Size = 9;
        //officersTable.Cell(7, 1).Range.Font.Size = 9;
        //officersTable.Cell(8, 1).Range.Font.Size = 9;
        //officersTable.Cell(9, 1).Range.Font.Size = 9;
        //officersTable.Cell(10, 1).Range.Font.Size = 9;
        //officersTable.Cell(11, 1).Range.Font.Size = 9;

        officersTable.Cell(1, 1).Range.Font.Bold = -1;
        officersTable.Cell(2, 1).Range.Font.Bold = -1;
        officersTable.Cell(3, 1).Range.Font.Bold = -1;
        officersTable.Cell(4, 1).Range.Font.Bold = -1;
        officersTable.Cell(5, 1).Range.Font.Bold = -1;
        officersTable.Cell(6, 1).Range.Font.Bold = -1;
        officersTable.Cell(7, 1).Range.Font.Bold = -1;
        officersTable.Rows[1].Range.Font.Size = 9;
        officersTable.Rows[2].Range.Font.Size = 9;
        officersTable.Rows[3].Range.Font.Size = 9;
        officersTable.Rows[4].Range.Font.Size = 9;
        officersTable.Rows[5].Range.Font.Size = 9;
        officersTable.Rows[6].Range.Font.Size = 9;
        officersTable.Rows[7].Range.Font.Size = 9;

        // get club officers for that club

        DataSet dsOfficers = new DataSet();

        SqlDataAdapter daOfficers = new SqlDataAdapter("select office, memberid, clubno from clubofficers where clubno = " + clubNo, conn);
        /*"and office not in ('CTREAS','CSAA','CSEC')",, conn);*/

        daOfficers.Fill(dsOfficers);

        DataTable officersDataTable = dsOfficers.Tables["Table"];
        officersTable.Cell(1, 1).Range.Text = "Pres";
        officersTable.Cell(2, 1).Range.Text = "VPE";
        officersTable.Cell(3, 1).Range.Text = "VPM";
        officersTable.Cell(4, 1).Range.Text = "VPPR";
        officersTable.Cell(5, 1).Range.Text = "Sec";
        officersTable.Cell(6, 1).Range.Text = "Treas";
        officersTable.Cell(7, 1).Range.Text = "SAA";

        foreach (DataRow rowOfficer in officersDataTable.Rows)
        {
          bool bFullMemberInfo = false;

          office = rowOfficer["Office"].ToString().Trim();
          memberID = (int)rowOfficer[1];
          email = "";
          name = "";
          loc1 = "";
          loc2 = "";
          phone1 = "";
          phone2 = "";

          GenerateMemberInfo(memberID, ref name, ref loc1, ref loc2, ref phone1, ref phone2, ref email, bSEC, bFullMemberInfo, true);
          string stuff = name;

          if (phone1.Length > 0)
            stuff += ", " + phone1;

          if (phone2.Length > 0)
            stuff += ", " + phone2;

          if (email.Length > 0)
            stuff += ", " + email;

          if (bSEC)
          {
            //if (division == "D")
            //  stuff = "Confidential";
            //else
            stuff = name;
          }



          if (office == "CPRES")
          {
            //Word.Range theRange = officersTable.Cell(5, 1).Range;
            //theRange.Text = "Pres" + "\t" + stuff;
            //theRange.Words[1].Font.Bold = -1;
            officersTable.Cell(1, 2).Range.Text = stuff;
          }
          else if (office == "CVPE")
          {
            //Word.Range theRange = officersTable.Cell(6, 1).Range;
            //theRange.Text = "VPE" + "\t" + stuff;
            //theRange.Words[1].Font.Bold = -1;
            officersTable.Cell(2, 2).Range.Text = stuff;
          }
          else if (office == "CVPM")
          {
            //Word.Range theRange = officersTable.Cell(7, 1).Range;
            //theRange.Text = "VPM" + "\t" + stuff;
            //theRange.Words[1].Font.Bold = -1;
            officersTable.Cell(3, 2).Range.Text = stuff;
          }
          else if (office == "CVPPR")
          {
            //Word.Range theRange = officersTable.Cell(8, 1).Range;
            //theRange.Text = "VPPR" + "\t" + stuff;
            //theRange.Words[1].Font.Bold = -1;
            officersTable.Cell(4, 2).Range.Text = stuff;
          }
          else if (office == "CSEC")
          {
            //Word.Range theRange = officersTable.Cell(9, 1).Range;
            //theRange.Text = "Sec" + "\t" + stuff;
            //theRange.Words[1].Font.Bold = -1;
            officersTable.Cell(5, 2).Range.Text = stuff;
          }
          else if (office == "CTREAS")
          {
            //Word.Range theRange = officersTable.Cell(10, 1).Range;
            //theRange.Text = "Treas" + "\t" + stuff;
            //theRange.Words[1].Font.Bold = -1;
            officersTable.Cell(6, 2).Range.Text = stuff;
          }
          else if (office == "CSAA")
          {
            //Word.Range theRange = officersTable.Cell(11, 1).Range;
            //theRange.Text = "SAA" + "\t" + stuff;
            //theRange.Words[1].Font.Bold = -1;
            officersTable.Cell(7, 2).Range.Text = stuff;
          }
        }
        /*
        clubLocationTable.Select();
        clubLocationTable.Rows.AllowBreakAcrossPages = 0;
        doc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = -1;

        officersTable.Select();
        officersTable.Rows.AllowBreakAcrossPages = 0;
        doc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = -1;

         */
        Word.Paragraph oPara4;
        oRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
        oPara4 = doc.Content.Paragraphs.Add(ref oRng);
        oPara4.Range.Font.Size = 4;
        oPara4.Format.SpaceAfter = 0;
      }
    }

    private void c3ToolStripMenuItem_Click(object sender, EventArgs e)
    {
      oWord = new Word.Application();
      SetUpDocument();
      GenerateAreaNew(oDoc, "C", 3);
    }

    private void subDirectoryToolStripMenuItem_Click(object sender, EventArgs e)
    {
      GenerateSubDirectory dlg = new GenerateSubDirectory();
      DialogResult res = dlg.ShowDialog();
      if (res == DialogResult.OK)
      {
        division = dlg.Division;
        String sArea = dlg.Area;
        //area = dlg.Area;

        area = System.Int32.Parse(sArea);
        if (division.Equals("E") && area.Equals(5))
          return;
        GenerateAreaThreadStart();
      }

      Word.Table tTable;

      int nTables = oDoc.Tables.Count;
      if (nTables < 1)
        return;

      for (int index = 1; index <= nTables; index++)
      {
        tTable = oDoc.Tables[index];
        tTable.Select();
        tTable.Rows.AllowBreakAcrossPages = 0;
        oDoc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = -1;

        //for (int rowIndex = 1; rowIndex < tTable.Rows.Count; rowIndex++)
        //  tTable.Rows[rowIndex].Range.ParagraphFormat.KeepWithNext = -1;
      }
    }

    private void areasToolStripMenuItem_Click(object sender, EventArgs e)
    {
      Thread divThread = new Thread(new ThreadStart(GenerateAreasThreadStart));
      divThread.Start();
    }

    void GenerateAreasThreadStart()
    {
      DataSet dsMatrix = new DataSet();
      SqlDataAdapter daMatrix = new SqlDataAdapter("Select * FROM DivAreaMatrix", conn);
      daMatrix.Fill(dsMatrix);
      DataTable dtMatrix = dsMatrix.Tables[0];
      //oWord = new Word.Application();

      Word._Application theWord = new Word.Application();
      foreach (DataRow row in dtMatrix.Rows)
      {
        String Division = row.ItemArray[1].ToString();
        object count = row.ItemArray[2];

        int numAreas = System.Convert.ToInt32(count);
        for (int area = 1; area <= numAreas; area++)
        {


          Word._Document oWordDoc = new Word.Document();
          //MAKING THE APPLICATION VISIBLE

          theWord.Visible = true;

          //ADDING A NEW DOCUMENT TO THE APPLICATION

          oWordDoc = theWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

          //SetUpDocument();
          //oWordDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

          oWordDoc.PageSetup.LineNumbering.Active = 0;
          oWordDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
          oWordDoc.PageSetup.TopMargin = theWord.InchesToPoints(.5F);
          oWordDoc.PageSetup.BottomMargin = 35;
          oWordDoc.PageSetup.LeftMargin = 35;
          oWordDoc.PageSetup.RightMargin = 25;
          oWordDoc.PageSetup.Gutter = 25;
          oWordDoc.PageSetup.HeaderDistance = 15;
          oWordDoc.PageSetup.FooterDistance = 15;
          //oDoc.PageSetup.PageWidth = 11;
          //oDoc.PageSetup.PageHeight = 8.5F;
          oWordDoc.PageSetup.FirstPageTray = Word.WdPaperTray.wdPrinterDefaultBin;
          oWordDoc.PageSetup.OtherPagesTray = Word.WdPaperTray.wdPrinterDefaultBin;
          oWordDoc.PageSetup.SectionStart = Word.WdSectionStart.wdSectionNewPage;
          oWordDoc.PageSetup.OddAndEvenPagesHeaderFooter = 0;
          oWordDoc.PageSetup.DifferentFirstPageHeaderFooter = 0;
          oWordDoc.PageSetup.VerticalAlignment = Word.WdVerticalAlignment.wdAlignVerticalTop;
          oWordDoc.PageSetup.SuppressEndnotes = 0;
          oWordDoc.PageSetup.MirrorMargins = 0;
          oWordDoc.PageSetup.TwoPagesOnOne = false;
          oWordDoc.PageSetup.BookFoldPrinting = true;
          //oDoc.PageSetup.BookFoldRevPrinting = true;
          //oDoc.PageSetup.BookFoldPrintingSheets = 1;
          oWordDoc.PageSetup.GutterPos = Word.WdGutterStyle.wdGutterPosLeft;

          GenerateAreaNew(oWordDoc, Division, area);
          Word.Table tTable;

          int nTables = oWordDoc.Tables.Count;
          if (nTables < 1)
            return;

          int t = oWordDoc.Paragraphs.Count;
          
          for (int index = 1; index <= nTables; index++)
          {
            tTable = oWordDoc.Tables[index];
            tTable.Select();
            tTable.Rows.AllowBreakAcrossPages = 0;
            oWordDoc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = -1;


            //for (int rowIndex = 1; rowIndex < tTable.Rows.Count; rowIndex++)
            //  tTable.Rows[rowIndex].Range.ParagraphFormat.KeepWithNext = -1;
          }
          //Object oSaveAsFile = (Object)"C:\\Area" + Division + area.ToString();// + ".doc";
          //oWordDoc.SaveAs(ref oSaveAsFile, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
          //  ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
          //Object oFalse = false;
          //oWordDoc.Close(ref oFalse, ref oMissing, ref oMissing);
        }
      }
    }

    public void GenerateClubs(/*Word._Document doc, string division, int area*/)
    {
      oWord = new Word.Application();
      SetUpDocument();
      // add clubs
      DataSet dsMatrix = new DataSet();
      SqlDataAdapter daMatrix = new SqlDataAdapter("Select * FROM DivAreaMatrix", conn);
      daMatrix.Fill(dsMatrix);
      DataTable dtMatrix = dsMatrix.Tables[0];

      foreach (DataRow row in dtMatrix.Rows)
      {
        String division = row.ItemArray[1].ToString();
        object count = row.ItemArray[2];

        int numAreas = System.Convert.ToInt32(count);
        Word.Table DivisionTitle;
        Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

        DivisionTitle = oDoc.Tables.Add(wrdRng, 1, 1, ref oMissing, ref oMissing);

        DivisionTitle.Range.ParagraphFormat.SpaceAfter = 0;
        DivisionTitle.Cell(1, 1).Range.Text = "Division " + division;// +area;
        DivisionTitle.Rows[1].Range.Font.Bold = 1;
        DivisionTitle.Rows[1].Range.Font.Size = 14;
        // areaGovTitle.Rows[2].Range.Font.Size = 8;
        DivisionTitle.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

        Word.Paragraph divBreakafter;
        object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
        divBreakafter = oDoc.Content.Paragraphs.Add(ref oRng);
        divBreakafter.Range.Font.Size = 4;
        divBreakafter.Format.SpaceAfter = 0;

        for (int area = 1; area <= numAreas; area++)
        {

          Word.Table areaTitle;
          wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

          areaTitle = oDoc.Tables.Add(wrdRng, 1, 1, ref oMissing, ref oMissing);

          areaTitle.Range.ParagraphFormat.SpaceAfter = 0;
          areaTitle.Cell(1, 1).Range.Text = "Area " + division +area;
          areaTitle.Rows[1].Range.Font.Bold = 1;
          areaTitle.Rows[1].Range.Font.Size = 12;
          // areaGovTitle.Rows[2].Range.Font.Size = 8;
          areaTitle.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
          areaTitle.Range.ParagraphFormat.SpaceAfter = 0;

          Word.Paragraph areaBreakafter;
          oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
          divBreakafter = oDoc.Content.Paragraphs.Add(ref oRng);
          divBreakafter.Range.Font.Size = 4;
          divBreakafter.Format.SpaceAfter = 0;

          // Clubs in Area
          DataSet dsClubs = new DataSet();
          SqlDataAdapter daClubs = new SqlDataAdapter("SELECT * FROM Clubs " +
              " WHERE Division = " + "'" + division + "'" + " AND Area = " + "'" + area + "'" +
               "Order by clubno", conn);

          daClubs.Fill(dsClubs);
          DataTable dataTableClub = dsClubs.Tables["Table"];
          string clubNo;
          string clubName;
          string dayOfTheWeek;
          string time;
          string web;
          string address;
          string frequency;
          string phone;
          string city;
          string zip;
          string email2;
          string clubStatus = "";

          foreach (DataRow rowClub in dataTableClub.Rows)
          {
            //object clubNumber = rowClub.ItemArray[0];

            // populate club info
            // add a table for the club information
            clubNo = rowClub.ItemArray[1].ToString().Trim();
            bool bSEC = false;

            //if (division == "A")
            //{
            //  if (clubNo == "2485" || clubNo == "3014")
            //    bSEC = true;
            //}

            //if (division == "C")
            //{
            //  if (clubNo == "6900" || clubNo == "7730")
            //    bSEC = true;
            //}

            if (division == "D")
            {
              if (clubNo == "782516" || clubNo == "1071907"  || clubNo == "9681" ||
                  clubNo == "1470790")
                bSEC = true;
            }

            if (division == "E")
            {
              if (clubNo == "7187" || clubNo == "8486" || clubNo == "771553" || clubNo == "814824" ||
                  clubNo == "722657" || clubNo == "1535248")
                bSEC = true;
            }


            clubName = rowClub.ItemArray[2].ToString().Trim();
            dayOfTheWeek = rowClub.ItemArray[8].ToString().Trim();
            time = rowClub.ItemArray[9].ToString().Trim();
            web = rowClub.ItemArray[18].ToString().Trim();
            phone = rowClub.ItemArray[14].ToString().Trim();
            string phone2 = rowClub.ItemArray[15].ToString().Trim();
            string email = rowClub.ItemArray[16].ToString().Trim();
            email2 = rowClub.ItemArray[17].ToString().Trim();
            string loc1 = rowClub.ItemArray[5].ToString().Trim();
            //loc2 = rowClub.ItemArray[9].ToString().Trim();
            address = rowClub.ItemArray[6].ToString().Trim();
            city = rowClub.ItemArray[10].ToString().Trim();
            zip = rowClub.ItemArray[11].ToString().Trim();
            frequency = rowClub.ItemArray[7].ToString().Trim();
            clubStatus = rowClub.ItemArray[20].ToString();

            Word.Table clubLocationTable;

            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            bool distinguishedClub = false;
            if (clubStatus.Contains("S") || clubStatus.Contains("P") || clubStatus.Contains("D"))
            {
              clubLocationTable = oDoc.Tables.Add(wrdRng, 4, 1, ref oMissing, ref oMissing);
              distinguishedClub = true;
            }
            else
              clubLocationTable = oDoc.Tables.Add(wrdRng, 3, 1, ref oMissing, ref oMissing);
            clubLocationTable.Range.ParagraphFormat.SpaceAfter = 0;
            //clubLocationTable.AllowAutoFit = true;
            clubLocationTable.BottomPadding = 0;

            // select table to keep from splitting across page breaks
            //int LastRow = clubLocationTable.Range.Rows.Count;
            //Word.Range MyRange = clubLocationTable.Range.Rows[0].Range;
            //MyRange.SetRange(0, LastRow);
            //MyRange = ActiveDocument.Tables(1).Range.Rows(2).Range
            //MyRange.SetRange Start:=MyRange.Start, _
            //End:=ActiveDocument.Tables(1).Range.Rows(LastRow).Range.End 
            //clubLocationTable.Range.SetRange
            //clubLocationTable.Select();
            //oDoc.ActiveWindow.Selection.ParagraphFormat.KeepWithNext = -1;

            clubLocationTable.AllowPageBreaks = false;
            //clubLocationTable.Selection.ParagraphFormat.KeepWithNext = true;
            //MyRange.ParagraphFormat.KeepWithNext = true;
            //areaGovTitle.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //clubLocationTable.Rows[1].Height = 15;
            string clubRow1;
            string clubRow2 = "";
            string clubRow3 = "";
            string clubRow4 = "";
            string clubRow5 = "";

            // need logic in club info, if something is empty we don't want a bunch of commas
            clubRow1 = "Club " + clubNo + " - " + clubName;
            //clubRow2 = dayOfTheWeek + " " + time + " " + web + " " + phone + " " + email;
            //if (dayOfTheWeek.Length > 1)
            //  clubRow2 += dayOfTheWeek + " ";

            if (frequency.Length > 1)
              clubRow2 += frequency + " ";

            if (time.Length > 1)
              clubRow2 += time + ", ";

            if (web.Length > 1)
              clubRow2 += web + ", ";

            if (phone.Length > 1)
              clubRow2 += phone + ", ";

            if (email.Length > 1)
              clubRow2 += email;

            //clubRow3 = loc1 + ", " + loc2 + ", " + address + ", " + city + ", " + zip;
            if (loc1.Length > 1)
              clubRow3 += loc1 + ", ";

            //if (loc2.Length > 1)
            //  clubRow3 += loc2 + ", ";

            if (address.Length > 1)
              clubRow3 += address + ", ";

            if (city.Length > 1)
              clubRow3 += city + ", ";

            if (zip.Length > 1)
              clubRow3 += zip;

            if (bSEC)
            {
              if (clubNo == "7187" || clubNo == "8486" || clubNo == "771553" || clubNo == "814824" ||
                  clubNo == "722657")
              {
                clubRow2 = "";
                if (dayOfTheWeek.Length > 1)
                  clubRow2 += dayOfTheWeek + " ";

                if (time.Length > 1)
                  clubRow2 += time + " ";

                clubRow3 = "Contact: Special Environment Clubs Chair, Randy Amelino, DTM";
                clubRow4 = "Email: chemdryall@gnww.net Phone: 951-258-6901";

                if (clubNo == "7187" || clubNo == "8486")
                  clubRow2 += "California Institute for Women, 16756 Chino-Corona Rd, Corona, 92880-9508";
                else if (clubNo == "771553" || clubNo == "814824")
                  clubRow2 += "California Rehab Center, Bldg 601 Men’s Education Dept, Norco, 92860";
                else if (clubNo == "722657")
                  clubRow2 += "California Rehab Center, Bldg 601 Men’s Education Dept, Norco, 92860";
              }

              if (clubNo == "782516" || clubNo == "1071907" || clubNo == "9681" || /*clubNo == "1167482" ||*/
                  clubNo == "1470790")
              {
                clubRow2 = "";
                if (dayOfTheWeek.Length > 1)
                  clubRow2 += dayOfTheWeek + " ";

                if (time.Length > 1)
                  clubRow2 += time + " ";

                clubRow3 = "";
                clubRow4 = "";
                clubRow5 = "";

                if (clubNo == "1167482")
                  clubRow2 += "Tameka Roberson, 760.922.0680x5074";
                else
                  clubRow2 += "Chuckwalla Staff, cindy.nepusz@cdcr.ca.gov, 760.922.5300x5124";


              }
            }

            clubLocationTable.Cell(1, 1).Range.Text = clubRow1;
            clubLocationTable.Rows[1].Range.Font.Bold = 1;
            clubLocationTable.Rows[1].Range.Font.Size = 11;

            if (distinguishedClub)
            {
              string clubStat = "2010-2011 ";
              switch (clubStatus)
              {
                case "P":
                  clubStat += "President's" + " ";
                  break;
                case "S":
                  clubStat += "Select" + " ";
                  break;
              }
              clubStat += "Distinguished Club";

              clubLocationTable.Cell(2, 1).Range.Text = clubStat;
              clubLocationTable.Cell(3, 1).Range.Text = clubRow2;
              clubLocationTable.Cell(4, 1).Range.Text = clubRow3;

              clubLocationTable.Rows[2].Range.Font.Bold = 1;
              clubLocationTable.Rows[2].Range.Font.Size = 9;
              clubLocationTable.Rows[3].Range.Font.Size = 9;
              clubLocationTable.Rows[4].Range.Font.Size = 9;
            }
            else
            {
              clubLocationTable.Cell(2, 1).Range.Text = clubRow2;
              clubLocationTable.Cell(3, 1).Range.Text = clubRow3;
              clubLocationTable.Rows[2].Range.Font.Size = 9;
              clubLocationTable.Rows[3].Range.Font.Size = 9;
            }

            Word.Paragraph clubLocBreakafter;

            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            clubLocBreakafter = oDoc.Content.Paragraphs.Add(ref oRng);
            clubLocBreakafter.Range.Font.Size = 4;
            clubLocBreakafter.Format.SpaceAfter = 0;
          }
        }

      }
    }

    private void clubHeadingsToolStripMenuItem_Click(object sender, EventArgs e)
    {

      GenerateClubs();
    }
  }
}
