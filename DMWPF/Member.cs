using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMWPF
{
  class Member
  {
    private int _clubNo;
    private int _memberID;
    private string _memberid;
    public string _lastName;
    public string _firstName;
    //public string _middleName;
    public string _title;
    //public string _workPhone;
    public string _homePhone;
    public string _cellPhone;
    public string _email;

    public Member(DataRow memberRow)
    {
      memberid = memberRow.ItemArray[0].ToString().Trim();
      FirstName = memberRow.ItemArray[1].ToString().Trim();
      LastName = memberRow.ItemArray[2].ToString().Trim();
      Title = memberRow.ItemArray[3].ToString().Trim();
      HomePhone = memberRow.ItemArray[4].ToString().Trim();
      CellPhone = memberRow.ItemArray[5].ToString().Trim();
      Email = memberRow.ItemArray[6].ToString().Trim();
    }

    public int ClubNo
    {
      get { return _clubNo; }
      set { _clubNo = value; }
    }

    public int MemberID
    {
      get { return _memberID; }
      set { _memberID = value; }
    }

    public string memberid
    {
      get { return _memberid; }
      set { _memberid = value; }
    }
    public string LastName
    {
      get { return _lastName; }
      set { _lastName = value; }
    }

    public string FirstName
    {
      get { return _firstName; }
      set { _firstName = value; }
    }

    public string Title
    {
      get { return _title; }
      set { _title = value; }
    }

    public string HomePhone
    {
      get { return _homePhone; }
      set { _homePhone = value; }
    }

    public string CellPhone
    {
      get { return _cellPhone; }
      set { _cellPhone = value; }
    }

    public string Email
    {
      get { return _email; }
      set { _email = value; }
    }
    //ClubNumber = System.Int32.Parse(rcd[0]);
    //MemberID = System.Int32.Parse(rcd[1]);
    //LastName = rcd[2];
    //FirstName = rcd[3];
    //MiddleName = rcd[4];
    //Title = rcd[5];
    //WorkPhone = rcd[6];
    //HomePhone = rcd[7];
    //CellPhone = rcd[8];
    //Email = rcd[9];
  }
}
