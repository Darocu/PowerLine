using System;
using System.Windows;
using System.Windows.Interop;

namespace CUSTPowerLineDialog
{
    /// <summary>
    /// Interaction logic for NoMCTPowerLineDialog.xaml
    /// </summary>
    public partial class NoMctPowerLineDialog : Window
    {
        public string SelectedReason { get; private set; }

        public NoMctPowerLineDialog()
        {
            InitializeComponent();
        }
        
        private void TitleBar_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed)
            {
                DragMove();
            }
        }
        
        private void Option_Click(object sender, RoutedEventArgs e)
        {
            if (OptionOther.IsChecked == true)
                TextOther.IsEnabled = true;
            else
            {
                TextOther.IsEnabled = false;
                TextOther.Text = string.Empty; // Clear when not in use
            }
        }
        
        private void RadioButton_GotKeyboardFocus(object sender, System.Windows.Input.KeyboardFocusChangedEventArgs e)
        {
            if (sender is System.Windows.Controls.RadioButton rb)
                rb.IsChecked = true;
        }
        
        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (!(Option1.IsChecked == true || Option2.IsChecked == true || Option3.IsChecked == true || OptionOther.IsChecked == true) ||
                (OptionOther.IsChecked == true && string.IsNullOrWhiteSpace(TextOther.Text.Trim())))
            {
                WarningText.Visibility = Visibility.Visible;
                return;
            }

            WarningText.Visibility = Visibility.Collapsed;

            if (OptionOther.IsChecked == true)
                SelectedReason = TextOther.Text.Trim();
            else if (Option1.IsChecked == true)
                SelectedReason = Option1.Content.ToString();
            else if (Option2.IsChecked == true)
                SelectedReason = Option2.Content.ToString();
            else if (Option3.IsChecked == true)
                SelectedReason = Option3.Content.ToString();

            TrySetDialogResult(true);
            Close();
        }
        
        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            TrySetDialogResult(false);
            Close();
        }

        private void TrySetDialogResult(bool? result)
        {
            try
            {
                // Only set DialogResult if window is modal
                if (this.IsLoaded && new WindowInteropHelper(this).Handle != IntPtr.Zero)
                    this.DialogResult = result;
            }
            catch (InvalidOperationException)
            {
                // Ignore if not shown as dialog
            }
        }
    }
}