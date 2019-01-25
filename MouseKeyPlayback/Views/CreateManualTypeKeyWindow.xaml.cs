using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace MouseKeyPlayback.Views
{
	/// <summary>
	/// Interaction logic for CreateManualTypeKeyWindow.xaml
	/// </summary>
	public partial class CreateManualTypeKeyWindow : Window
	{
		public string text { get; set; }

		public CreateManualTypeKeyWindow()
		{
			InitializeComponent();

			System.Windows.Application curApp = System.Windows.Application.Current;
			Window mainWindow = curApp.MainWindow;
			this.Left = mainWindow.Left + (mainWindow.Width - this.ActualWidth) / 2;
			this.Top = mainWindow.Top + (mainWindow.Height - this.ActualHeight) / 2;
		}

		private void BtnCancel_Click(object sender, RoutedEventArgs e)
		{
			this.Close();
		}

		private void BtnOk_Click(object sender, RoutedEventArgs e)
		{
			text = tbxInsert.Text;
			this.Close();
		}
	}
}
