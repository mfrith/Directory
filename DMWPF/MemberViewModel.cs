using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace DMWPF
{
  class MemberViewModel : PropertyChangedBase
  {
    private int _clubNo = -1;
    private int _memberID = -1;
    public string _lastName = string.Empty;
    public string _firstName = string.Empty;
    //public string _middleName;
    public string _title = string.Empty;
    //public string _workPhone;
    public string _homePhone = string.Empty;
    public string _cellPhone = string.Empty;
    public string _email = string.Empty;

    private Member _member = null;

    public MemberViewModel()
    {

    }

    public MemberViewModel(Member member)
    {
      _member = member;

    }
    public int ClubNo
    {
      get { return _clubNo; }
      set { SetProperty(ref _clubNo, value, () => ClubNo); }
    }

    public int MemberID
    {
      get { return _memberID; }
      set { SetProperty(ref _memberID, value, () => MemberID); }
    }

    public string LastName
    {
      get { return _lastName; }
      set { SetProperty(ref _lastName, value, () => LastName); }
    }

    public string FirstName
    {
      get { return _firstName; }
      set { SetProperty(ref _firstName, value, () => FirstName); }
    }

    public string Title
    {
      get { return _title; }
      set { SetProperty(ref _title, value, () => Title); }
    }

    public string HomePhone
    {
      get { return _homePhone; }
      set { SetProperty(ref _homePhone, value, () => HomePhone); }
    }

    public string CellPhone
    {
      get { return _cellPhone; }
      set { SetProperty(ref _cellPhone, value, () => CellPhone); }
    }

    public string Email
    {
      get { return _email; }
      set { SetProperty(ref _email, value, () => Email); }
    }
  }
}
