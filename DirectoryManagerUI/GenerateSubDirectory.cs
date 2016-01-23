using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DistrictManager
{
  public partial class GenerateSubDirectory : Form
  {
    public string division = "";
    public string area = "";

    public GenerateSubDirectory()
    {
      InitializeComponent();
    }

    public string Division
    {
      get
      {
        return division;
      }

      set
      {
        DivisionTextBox.Text = value;
      }
    }

    public string Area
    {
      get
      { return area; }

    }
    private void DivComboBox_SelectedIndexChanged(object sender, EventArgs e)
    {
      object selection = DivComboBox.Items[DivComboBox.SelectedIndex];
      division = selection.ToString();

    }

    private void AreaComboBox_SelectedIndexChanged(object sender, EventArgs e)
    {
      object selection = AreaComboBox.Items[AreaComboBox.SelectedIndex];
      area = selection.ToString();
    }

    private void DivisionTextBox_TextChanged(object sender, EventArgs e)
    {
      division = sender.ToString();
    }

    private void AreaTextBox_TextChanged(object sender, EventArgs e)
    {
      //string txtArea = sender.ToString();
      //if (txtArea == "1")
      //  area = 1;
      //else if (txtArea == "2")
      //  area = 2;
      //else if (txtArea == "3")
      //  area = 3;
      //else if (txtArea == "4")
      //  area = 4;
      //else if (txtArea == "5")
      //  area = 5;
      //sender.Text.ToString();
    }

    private void buttonOk_Click(object sender, EventArgs e)
    {
      this.DialogResult = DialogResult.OK;
      this.Close();
    }

  }
}
