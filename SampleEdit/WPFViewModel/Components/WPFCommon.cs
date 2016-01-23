using System;
using System.Data;
using System.Windows.Controls;
using System.Windows.Data;

namespace Common.Library
{
  public class WPFCommon
  {
    public static GridView CreateGridViewColumns(DataTable dt)
    {
      // Create the GridView
      GridView gv = new GridView();
      gv.AllowsColumnReorder = true;

      // Create the GridView Columns
      foreach (DataColumn item in dt.Columns)
      {
        GridViewColumn gvc = new GridViewColumn();
        gvc.DisplayMemberBinding = new Binding(item.ColumnName);
        gvc.Header = item.ColumnName;
        gvc.Width = Double.NaN;
        gv.Columns.Add(gvc);
      }

      return gv;
    }
  }
}
