using System.Windows;
using System.Windows.Controls;

namespace WPFViewModel
{
  public partial class MainWindow : Window
  {
    private ShirtViewModel _ViewModel = null;

    public MainWindow()
    {
      InitializeComponent();

      _ViewModel = (ShirtViewModel)this.Resources["viewModel"];
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      _ViewModel.LoadAll();
    }

    private void btnAdd_Click(object sender, RoutedEventArgs e)
    {
      _ViewModel.SetAddMode();
    }

    private void btnSave_Click(object sender, RoutedEventArgs e)
    {
      _ViewModel.Save();
    }

    private void btnReload_Click(object sender, RoutedEventArgs e)
    {
      _ViewModel.Reload();
    }

    private void btnSearch_Click(object sender, RoutedEventArgs e)
    {
      _ViewModel.Search();
    }

    private void btnCancel_Click(object sender, RoutedEventArgs e)
    {
      _ViewModel.Cancel();
    }

    private void btnDelete_Click(object sender, RoutedEventArgs e)
    {
      if (MessageBox.Show("Delete this Officer?", "Delete?", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
        _ViewModel.DeleteAnOfficer();
    }

    private void TextChanged(object sender, TextChangedEventArgs e)
    {
      if (((UIElement)sender).IsKeyboardFocused)  // Only Change Mode if Element has Keyboard Focus
        _ViewModel.EditMode();
    }

    private void SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
      if (((UIElement)sender).IsKeyboardFocused || ((UIElement)sender).IsMouseDirectlyOver)
        _ViewModel.EditMode();
    }

    private void lstData_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {

    }
  }
}
