using System.Configuration;

namespace Common.Library
{
  public class AppSettings
  {
    public static string ConnectString
    {
      get { return ConfigurationManager.ConnectionStrings["ShirtSample"].ConnectionString; }
    }
  }
}
