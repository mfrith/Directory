using System;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.SqlClient;
using System.Linq.Expressions;
using System.Linq;
using Common.Library;
using System.Collections.Generic;

namespace WPFViewModel
{
  public class ShirtViewModel : CommonBase
  {
    #region Constructor
    public ShirtViewModel()
    {
      NormalMode();
      Init();
    }
    #endregion

    #region Private Variables
    private bool _IsAddMode = false;
    private bool _IsAddVisible = true;
    private bool _IsDeleteVisible = true;
    private bool _IsSaveVisible = false;
    private bool _IsCancelVisible = false;
    private bool _IsValidationVisible = false;
    private bool _IsListEnabled = true;

    private ObservableCollection<ValidationMessage> _Messages = new ObservableCollection<ValidationMessage>();

    private int _SelectedIndex = -1;
    private int _ClubOfficerId = -1;
    private int _MemberId = -1;
    private int _SearchMemberId = -1;
    private int _ClubNo = 0;
    private string _Office = string.Empty;
    private string _LastName = string.Empty;
    private string _SearchText = string.Empty;
    private string _title = string.Empty;

    private DataTable _Members;
    private DataTable _ClubOfficers;
    private DataTable _SearchResults;


    #region Public Properties
    public bool IsAddMode
    {
      get { return _IsAddMode; }
      set
      {
        if (_IsAddMode != value)
        {
          _IsAddMode = value;
          RaisePropertyChanged("IsAddMode");
        }
      }
    }

    public bool IsAddVisible
    {
      get { return _IsAddVisible; }
      set
      {
        if (_IsAddVisible != value)
        {
          _IsAddVisible = value;
          RaisePropertyChanged("IsAddVisible");
        }
      }
    }

    public bool IsDeleteVisible
    {
      get { return _IsDeleteVisible; }
      set
      {
        if (_IsDeleteVisible != value)
        {
          _IsDeleteVisible = value;
          RaisePropertyChanged("IsDeleteVisible");
        }
      }
    }

    public bool IsSaveVisible
    {
      get { return _IsSaveVisible; }
      set
      {
        if (_IsSaveVisible != value)
        {
          _IsSaveVisible = value;
          RaisePropertyChanged("IsSaveVisible");
        }
      }
    }

    public bool IsCancelVisible
    {
      get { return _IsCancelVisible; }
      set
      {
        if (_IsCancelVisible != value)
        {
          _IsCancelVisible = value;
          RaisePropertyChanged("IsCancelVisible");
        }
      }
    }

    public bool IsValidationVisible
    {
      get { return _IsValidationVisible; }
      set
      {
        if (_IsValidationVisible != value)
        {
          _IsValidationVisible = value;
          RaisePropertyChanged("IsValidationVisible");
        }
      }
    }

    public bool IsListEnabled
    {
      get { return _IsListEnabled; }
      set
      {
        if (_IsListEnabled != value)
        {
          _IsListEnabled = value;
          RaisePropertyChanged("IsListEnabled");
        }
      }
    }

    public ObservableCollection<ValidationMessage> Messages
    {
      get { return _Messages; }
      set
      {
        if (_Messages != value)
        {
          _Messages = value;
          RaisePropertyChanged("Messages");
        }
      }
    }
    
    public int ClubNo
    {
      get { return _ClubNo; }
      set
      {
        if (_ClubNo != value)
        {
          _ClubNo = value;
          RaisePropertyChanged("ClubNo");
        }
      }
    }

    public string Office
    {
      get { return _Office; }
      set
      {
        if (_Office != value)
        {
          _Office = value;
          RaisePropertyChanged("Office");
        }
      }
    }
    
    public string LastName
    {
      get { return _LastName; }
      set
      {
        if (_LastName != value)
        {
          _LastName = value;
          RaisePropertyChanged("LastName");
        }
      }
    }
    
    public string SearchText
    {
      get { return _SearchText; }
      set
      {
        if (_SearchText != value)
        {
          _SearchText = value;
          RaisePropertyChanged("SearchText");
        }
      }
    }

    public int ClubOfficerId
    {
      get { return _ClubOfficerId; }
      set
      {
        if (_ClubOfficerId != value)
        {
          _ClubOfficerId = value;
        }
      }
    }

    public int SelectedIndex
    {
      get { return _SelectedIndex; }
      set
      {
        if (_SelectedIndex != value)
        {
          _SelectedIndex = value;
          DisplayAShirt();
          RaisePropertyChanged("SelectedIndex");
        }
      }
    }
    
    public int MemberId
    {
      get { return _MemberId; }
      set
      {
        if (_MemberId != value)
        {
          _MemberId = value;
          RaisePropertyChanged("MemberId");
        }
      }
    }

    public int SearchMemberId
    {
      get { return _SearchMemberId; }
      set
      {
        if (_SearchMemberId != value)
        {
          _SearchMemberId = value;
          RaisePropertyChanged("SearchMemberId");
        }
      }
    }

    public DataTable ClubOfficers
    {
      get { return _ClubOfficers; }
      set
      {
        if (_ClubOfficers != value)
        {
          _ClubOfficers = value;
          SelectedIndex = 0;
          RaisePropertyChanged("ClubOfficers");
        }
      }
    }

    public DataTable SearchResults
    {
      get { return _SearchResults; }
      set
      {
        if (_SearchResults != value)
        {
          _SearchResults = value;
          RaisePropertyChanged("SearchResults");
        }
      }
    }

    public DataTable Members
    {
      get { return _Members; }
      set
      {
        if (_Members != value)
        {
          _Members = value;
          SelectedIndex = 0;
          RaisePropertyChanged("Members");
        }
      }
    }
    #endregion Public Properties

    #region Init Method
    public void Init()
    {
      IsValidationVisible = false;
      IsListEnabled = true;

      Messages.Clear();

      ClubOfficerId = -1;

    }
    #endregion

    #region LoadAll
    public void LoadAll()
    {
      LoadClubOfficers();
      LoadMembers(); 
    }
    #endregion

    public void LoadClubOfficers()
    {
      DataTable dt = null;

      try
      {
        dt = DataLayer.GetDataTable("SELECT * FROM Club_Officer_View order by ClubNo",
          AppSettings.ConnectString);

        ClubOfficers = dt;

        NormalMode();
      }
      catch (Exception ex)
      {
        DisplayMessages(ex.Message);
      }
    }

    public void Reload()
    {
      LoadMembers();
    }

    #region LoadColors Method
    public void LoadMembers()
    {
      DataTable dt = null;

      try
      {
        dt = DataLayer.GetDataTable("SELECT MemberId FROM Members",
          AppSettings.ConnectString);

        Members = dt;
      }
      catch (Exception ex)
      {
        DisplayMessages(ex.Message);
      }
    }
    #endregion

    #region DisplayAShirt Method
    /// <summary>
    /// This method is called from the SelectedIndex property set.
    /// </summary>
    public void DisplayAShirt()
    {
      string sql;
      SqlCommand cmd;
      DataTable dt = null;

      if (SelectedIndex != -1)
      {
        // Get Index from Shirts Collection
        ClubOfficerId = Convert.ToInt32(ClubOfficers.Rows[SelectedIndex]["ClubOfficerId"]);

        sql = "SELECT * FROM ClubOfficers WHERE ClubOfficerId = @ClubOfficerId";
        try
        {
          cmd = new SqlCommand(sql);
          cmd.Parameters.Add(new SqlParameter("@ClubOfficerId", ClubOfficerId));
          cmd.Connection = new SqlConnection(AppSettings.ConnectString);

          dt = DataLayer.GetDataTable(cmd);
          if (dt.Rows.Count > 0)
          {
            //ShirtName = dt.Rows[0]["ShirtName"].ToString();
            Office = Convert.ToString(dt.Rows[0]["Office"]);
            ClubNo = Convert.ToInt32(dt.Rows[0]["ClubNo"]);
            MemberId = Convert.ToInt32(dt.Rows[0]["MemberId"]);
          }
        }
        catch (Exception ex)
        {
          DisplayMessages(ex.Message);
        }
      }
    }
    #endregion

    #region DeleteAnOfficer Method
    public bool DeleteAnOfficer()
    {
      bool ret = false;
      string sql;
      SqlCommand cmd;
      int rows = 0;

      sql = "DELETE FROM ClubOfficers WHERE ClubOfficerId = @ClubOfficerId";
      try
      {
        cmd = new SqlCommand(sql);
        cmd.Parameters.Add(new SqlParameter("@ClubOfficerId", ClubOfficerId));
        cmd.Connection = new SqlConnection(AppSettings.ConnectString);

        rows = DataLayer.ExecuteSQL(cmd);

        if (rows == 1)
          ret = true;

        if (ret)
          // Redisplay all Shirts
          LoadClubOfficers();
        else
          DisplayMessages("Can't find Officer to Delete");
      }
      catch (Exception ex)
      {
        DisplayMessages(ex.Message);
      }

      return ret;
    }  
    #endregion
 
    #region SetAddMode Method
    public void SetAddMode()
    {
      IsAddMode = true;
      Office = string.Empty;

      EditMode();
    } 
    #endregion

    #region Search Method
    public void Search()
    {
      bool ret = false;
      string sql;
      SqlCommand cmd;
      int rows = 0;

      DataTable dt = null;
      List<string> searchNames = SearchText.Split(',').ToList();
      string c = string.Join(",", searchNames.Select(x => "\'" + x.Trim() + "\'"));

      sql = "SELECT FirstName, LastName, Title, MemberId from Members WHERE (Lastname in (" + c + "))";


      try
      {
        //cmd = new SqlCommand(sql);
        //cmd.Parameters.Add(new SqlParameter("@SearchText", SearchText));
        //cmd.Connection = new SqlConnection(AppSettings.ConnectString);

        dt = DataLayer.GetDataTable(sql, AppSettings.ConnectString);
        SearchResults = dt;
      }
      catch (Exception ex)
      {
        DisplayMessages(ex.Message);
      }              
    }
    #endregion

    #region Save Method
    public void Save()
    {
      bool success = false;

      if (IsAddMode)
      {
        success = InsertData();
      }
      else
      {
        success = UpdateData();
      }

      if (success)
        NormalMode();
    }
    #endregion

    #region Cancel Method
    public void Cancel()
    {
      DisplayAShirt();

      NormalMode();
    }
    #endregion

    #region DataValidate Method
    private bool DataValidate()
    {
      bool ret = false;

      Messages.Clear();

      if (MemberId == -1)
        Messages.Add(new ValidationMessage("record must be selected."));

      ret = (Messages.Count == 0);

      if (!ret)
      {
        IsValidationVisible = true;
      }

      return ret;
    }
    #endregion
 
    #region InsertData Method
    private bool InsertData()
    {
      bool ret = false;
      string sql;
      SqlCommand cmd;
      int rows = 0;

      sql = "INSERT INTO ClubOfficers(ClubNo, Office, MemberId) ";
      sql += " VALUES(@ClubNo, @Office, @MemberId) ";

      try
      {
        cmd = new SqlCommand(sql);
        cmd.Connection = new SqlConnection(AppSettings.ConnectString);
        cmd.Parameters.Add(new SqlParameter("@ClubNo", Convert.ToInt32(ClubNo)));
        cmd.Parameters.Add(new SqlParameter("@Office", Office));
        cmd.Parameters.Add(new SqlParameter("@MemberId", Convert.ToInt32(MemberId)));

        rows = DataLayer.ExecuteSQL(cmd);

        ret = (rows == 1);

        // Reload All Shirts
        if (ret)
        {
          IsAddMode = false;
          LoadClubOfficers();
        }
      }
      catch (Exception ex)
      {
        DisplayMessages(ex.Message);
      }

      return ret;
    }    
    #endregion
    
    #region UpdateData Method
    private bool UpdateData()
    {
      bool ret = false;
      string sql;
      SqlCommand cmd;
      int rows = 0;

      sql = "UPDATE ClubOfficers SET ";
      sql += " MemberId = @MemberId, ";
      sql += " Office = @Office, ";
      sql += " ClubNo = @ClubNo ";
      sql += " WHERE ClubOfficerId = @ClubOfficerId ";

      try
      {
        cmd = new SqlCommand(sql);
        cmd.Connection = new SqlConnection(AppSettings.ConnectString);
        cmd.Parameters.Add(new SqlParameter("@MemberId", MemberId));
        cmd.Parameters.Add(new SqlParameter("@Office", Office));
        cmd.Parameters.Add(new SqlParameter("@ClubNo", ClubNo));
        cmd.Parameters.Add(new SqlParameter("@ClubOfficerId", ClubOfficerId));

        rows = DataLayer.ExecuteSQL(cmd);

        ret = (rows == 1);

        // Reload All Shirts
        if (ret)
          LoadClubOfficers();
        else
          DisplayMessages("Can't Find Shirt Id: " + ClubOfficerId.ToString() + " to update it.");
      }
      catch (Exception ex)
      {
        DisplayMessages(ex.Message);
      }

      return ret;
    }
    #endregion

    #region UI State Modes
    public void EditMode()
    {
      IsAddVisible = false;
      IsDeleteVisible = false;
      IsSaveVisible = true;
      IsCancelVisible = true;
      IsListEnabled = false;
    }

    private void NormalMode()
    {
      IsAddVisible = true;
      IsDeleteVisible = true;
      IsSaveVisible = false;
      IsCancelVisible = false;
      IsValidationVisible = false;
      IsListEnabled = true;
      Messages.Clear();
    }
    #endregion

    #region DisplayMessages
    private void DisplayMessages(string msg)
    {
      IsValidationVisible = true;
      Messages.Add(new ValidationMessage(msg));
    }
    #endregion
  }
  #endregion
}