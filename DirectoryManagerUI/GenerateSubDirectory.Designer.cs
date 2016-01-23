namespace DistrictManager
{
  partial class GenerateSubDirectory
  {
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
      if (disposing && (components != null))
      {
        components.Dispose();
      }
      base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
      this.DivComboBox = new System.Windows.Forms.ComboBox();
      this.AreaComboBox = new System.Windows.Forms.ComboBox();
      this.buttonOk = new System.Windows.Forms.Button();
      this.DivisionTextBox = new System.Windows.Forms.TextBox();
      this.AreaTextBox = new System.Windows.Forms.TextBox();
      this.SuspendLayout();
      // 
      // DivComboBox
      // 
      this.DivComboBox.FormattingEnabled = true;
      this.DivComboBox.Items.AddRange(new object[] {
            "A",
            "B",
            "C",
            "D",
            "E"});
      this.DivComboBox.Location = new System.Drawing.Point(29, 40);
      this.DivComboBox.Name = "DivComboBox";
      this.DivComboBox.Size = new System.Drawing.Size(121, 21);
      this.DivComboBox.Sorted = true;
      this.DivComboBox.TabIndex = 0;
      this.DivComboBox.SelectedIndexChanged += new System.EventHandler(this.DivComboBox_SelectedIndexChanged);
      // 
      // AreaComboBox
      // 
      this.AreaComboBox.FormattingEnabled = true;
      this.AreaComboBox.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6"});
      this.AreaComboBox.Location = new System.Drawing.Point(183, 40);
      this.AreaComboBox.Name = "AreaComboBox";
      this.AreaComboBox.Size = new System.Drawing.Size(121, 21);
      this.AreaComboBox.TabIndex = 1;
      this.AreaComboBox.SelectedIndexChanged += new System.EventHandler(this.AreaComboBox_SelectedIndexChanged);
      // 
      // buttonOk
      // 
      this.buttonOk.Location = new System.Drawing.Point(229, 83);
      this.buttonOk.Name = "buttonOk";
      this.buttonOk.Size = new System.Drawing.Size(75, 23);
      this.buttonOk.TabIndex = 2;
      this.buttonOk.Text = "Ok";
      this.buttonOk.UseVisualStyleBackColor = true;
      this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
      // 
      // DivisionTextBox
      // 
      this.DivisionTextBox.Location = new System.Drawing.Point(50, 12);
      this.DivisionTextBox.Name = "DivisionTextBox";
      this.DivisionTextBox.Size = new System.Drawing.Size(100, 20);
      this.DivisionTextBox.TabIndex = 3;
      this.DivisionTextBox.Visible = false;
      this.DivisionTextBox.TextChanged += new System.EventHandler(this.DivisionTextBox_TextChanged);
      // 
      // AreaTextBox
      // 
      this.AreaTextBox.Location = new System.Drawing.Point(204, 12);
      this.AreaTextBox.Name = "AreaTextBox";
      this.AreaTextBox.Size = new System.Drawing.Size(100, 20);
      this.AreaTextBox.TabIndex = 4;
      this.AreaTextBox.Visible = false;
      this.AreaTextBox.TextChanged += new System.EventHandler(this.AreaTextBox_TextChanged);
      // 
      // GenerateSubDirectory
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(331, 138);
      this.Controls.Add(this.AreaTextBox);
      this.Controls.Add(this.DivisionTextBox);
      this.Controls.Add(this.buttonOk);
      this.Controls.Add(this.AreaComboBox);
      this.Controls.Add(this.DivComboBox);
      this.Name = "GenerateSubDirectory";
      this.Text = "GenerateSubDirectory";
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion

    private System.Windows.Forms.ComboBox DivComboBox;
    private System.Windows.Forms.ComboBox AreaComboBox;
    private System.Windows.Forms.Button buttonOk;
    private System.Windows.Forms.TextBox DivisionTextBox;
    private System.Windows.Forms.TextBox AreaTextBox;
  }
}