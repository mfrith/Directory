using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DMWPF
{
  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : Window
  {
    
    DirectoryInfo _di = new DirectoryInfo();

    public MainWindow()
    {
      InitializeComponent();

    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      //Task t = await _di.LoadMembersAsync();
      //MemberGrid.DataContext = _di.LoadMembersAsync();

      _di.LoadMembers();
      MemberGrid.DataContext = _di;
      DirectoryTree.DataContext = _di;
      //DirectoryTree.ItemsSource = _di.Divisions;
      TreeViewItem tvi = new TreeViewItem();
      //tvi.Name = "District 12";
      
      //DirectoryTree.set
    }
  }
}
