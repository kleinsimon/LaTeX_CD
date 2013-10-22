using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
    /// Interaktionslogik für errorLog.xaml
    /// </summary>
    public partial class errorLog : Window
    {
        public string DisplayString
        {
            get { return OutputText.Text; }
            set { OutputText.Text = value; }
        }

        public errorLog()
        {
            InitializeComponent();
        }

        public errorLog(string DisplayString)
        {
            InitializeComponent();
            this.DisplayString = DisplayString;
            this.Show();
        }
    }
}
