using System.Data;
using System.Data.SqlClient;

namespace Common.Library
{
  /// <summary>
  /// A simple data layer for sample applications
	/// We assume that you will be using EF, LINQ to SQL, or something else for your data layer
  /// </summary>
  public class DataLayer
  {
    #region GetDataSet Methods
    public static DataSet GetDataSet(string sql, string connectString)
    {
      SqlCommand cmd = null;
      SqlConnection cnn = null;

      //  Create Command Object
      cmd = new SqlCommand(sql);
      //  Create Connection Object
      cnn = new SqlConnection(connectString);
      //  Assign Connection to Command Object
      cmd.Connection = cnn;

      //  Call Overloaded GetDataSet method
      return GetDataSet(cmd);
    }

    public static DataSet GetDataSet(SqlCommand cmd)
    {
      DataSet ds = new DataSet();
      SqlDataAdapter da = null;

      //  Create Data Adapter
      da = new System.Data.SqlClient.SqlDataAdapter(cmd);
      da.Fill(ds);
      cmd.Connection.Close();
      cmd.Connection.Dispose();

      return ds;
    }
    #endregion

    #region GetDataTable Methods
    public static DataTable GetDataTable(string sql, string connectString)
    {
      return GetDataSet(sql, connectString).Tables[0];
    }

    public static DataTable GetDataTable(SqlCommand cmd)
    {
      return GetDataSet(cmd).Tables[0];
    }
    #endregion

    #region ExecuteSQL Methods
    public static int ExecuteSQL(IDbCommand cmd)
    {
      int ret = 0;

      if (cmd.Connection.State == ConnectionState.Closed)
        cmd.Connection.Open();

      ret = cmd.ExecuteNonQuery();
      cmd.Connection.Close();
      cmd.Connection.Dispose();

      return ret;
    }

    public static int ExecuteSQL(string sql, string connectString)
    {
      SqlCommand cmd = null;
      int ret = 0;

      //  Create Command Object
      cmd = new SqlCommand(sql);
      // Create Connection Object
      cmd.Connection = new SqlConnection(connectString);
      // Open Connection
      cmd.Connection.Open();

      //  Execute SQL
      ret = ExecuteSQL(cmd);

      // Close & Dispose of Connection
      cmd.Connection.Close();
      cmd.Connection.Dispose();

      return ret;
    }
    #endregion

    #region 'Create' methods
    public static SqlCommand CreateCommand(string sql)
    {
      return new SqlCommand(sql);
    }

    public static SqlConnection CreateConnection(string connectString)
    {
      return new SqlConnection(connectString);
    }

    public static SqlParameter CreateParameter(string name, object value)
    {
      return new SqlParameter(name, value);
    }
    #endregion
  }
}