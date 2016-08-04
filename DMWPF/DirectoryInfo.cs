using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMWPF
{
  class Division
  {
    private string _name;
    public string Name
    {
      get { return _name; }
      set { _name = value; }
    }

  }
  class DirectoryInfo
  {
    SqlConnection conn = new SqlConnection();

    ObservableCollection<MemberViewModel> _memberVM = new ObservableCollection<MemberViewModel>();
    ObservableCollection<Member> _members = new ObservableCollection<Member>();
    internal Task LoadMembersAsync()
    {
      
      return Task.FromResult(0);
    }

    internal void LoadMembers()
    {
      conn.ConnectionString = @"Server=.\SQLEXPRESS;Database=D12;Integrated Security=true;";
      conn.Open(); 

      DataSet dsMembers = new DataSet();
      SqlDataAdapter daMember = new SqlDataAdapter();
      SqlCommand clubOfficerCmd = new SqlCommand("select memberid, firstname, lastname, title, " +
        "homephone, cellphone, email from members order by memberid", conn);

      daMember.SelectCommand = clubOfficerCmd;

      daMember.Fill(dsMembers);
      DataTable dtMembers = dsMembers.Tables[0];
      if (dtMembers.Rows.Count < 1)
        return;

      DataRow rowMember;
      for (int i = 0; i < dtMembers.Rows.Count; i++ )
      {
        rowMember = dtMembers.Rows[i];
        //MemberViewModel m = new MemberViewModel(new Member(rowMember));
        Member m = new Member(rowMember);
        _members.Add(m);
      }

    }

    public ObservableCollection<Member> Members
    {
      get { return _members; }
    }

    private ObservableCollection<Division> _divisions = new ObservableCollection<Division>();
    public ObservableCollection<Division> Divisions
    {
      get { return _divisions; }
    }
    //public ObservableCollection<MemberViewModel> Members
    //{
    //  get { return _memberVM; }
    //}
  }
}
