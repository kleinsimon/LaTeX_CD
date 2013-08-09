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

namespace Latex2CD2
{
    /// <summary>
    /// Interaktionslogik für TemplateWindow.xaml
    /// </summary>
    public partial class TemplateWindow : Window
    {
        public TemplateWindow()
        {
            InitializeComponent();
            TemplateText.Text = (Properties.Settings.Default.LastTemplate == string.Empty) ? Properties.Settings.Default.DefaultTemplate : Properties.Settings.Default.LastTemplate;
        }

        private void DefaultButton_Click(object sender, RoutedEventArgs e)
        {
            TemplateText.Text = Properties.Settings.Default.DefaultTemplate;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.LastTemplate = TemplateText.Text;

            Properties.Settings.Default.Save();
            this.Close();
        }
    }
}
