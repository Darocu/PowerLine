using System.Windows;

namespace CUSTPowerLineDialog
{
    public partial class MainWindow : Window
    {
        public MainWindow()
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

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void OpenDialog_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new NoMctPowerLineDialog();
            var result = dialog.ShowDialog();
            if (result == true)
            {
                MessageBox.Show($"Selected Reason: {dialog.SelectedReason}", "Dialog Result");
            }
        }
    }
}