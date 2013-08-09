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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Threading;
using System.Diagnostics;
using System.IO;
using CorelDRAW;
using VGCore;

namespace Latex2CD2
{
    enum ShapeProperties
    {
        isLatexObject = 1,
        latexText = 2,
        latexTemplate = 3,
    }

    /// <summary>
    /// Interaktionslogik für UserControl1.xaml
    /// </summary>
    public partial class LatexDock : UserControl
    {
        CorelDRAW.Application CDWin;
        CorelDRAW.Shape curShape;
        bool locked = false;
        string curTemplate = "";
        string tmpfp = System.IO.Path.GetTempPath();
        string tmpf = "";

        public LatexDock()
        {
            InitializeComponent();
            getTemplate();
        }

        private void getTemplate()
        {
            if (!Properties.Settings.Default.LastTemplate.Contains("%%ANCHOR%%"))
                curTemplate = Properties.Settings.Default.DefaultTemplate;
            else
                curTemplate = Properties.Settings.Default.LastTemplate;
        }

        public LatexDock(object app)
        {
            InitializeComponent();
            getTemplate();

            CDWin = (CorelDRAW.Application)app;
            CDWin.SelectionChange += CDWin_SelectionChange;
            getSelection();
        }

        void CDWin_SelectionChange()
        {
            if (locked) return;
            InsertButton.Content = "Insert";
            OutputText.Text = "";

            getTemplate();

            getSelection();
        }

        private void getSelection()
        {
            if (locked) return;

            try
            {
                curShape = CDWin.Application.ActiveShape;
                if (curShape.Properties["latex2cd", (int)ShapeProperties.isLatexObject] == true)
                {
                    editShape(curShape);
                    InsertButton.Content = "Save";
                }
                else
                {
                    curShape = null;
                }
            }
            catch
            {
                return;
            }
        }

        private void editShape(CorelDRAW.Shape curShape)
        {
            OutputText.Text = curShape.Properties["latex2cd", (int)ShapeProperties.latexText];
            curTemplate = curShape.Properties["latex2cd", (int)ShapeProperties.latexTemplate];
        }

        private void InsertButton_Click(object sender, RoutedEventArgs e)
        {
            GeneratePDF(curShape);
            cleanTempFiles();
        }

        private void GeneratePDF(CorelDRAW.Shape oldShape = null)
        {
            if (!ExistsOnPath(Properties.Settings.Default.PDFLatexPath))
            {
                MessageBox.Show("pdflatex.exe not in PATH, exiting... Maybe install MikTex?");
                return;
            }
            if (!ExistsOnPath(Properties.Settings.Default.PDFLatexPath))
            {
                MessageBox.Show("mgs.exe not in PATH, exiting... Maybe install MikTex?");
                return;
            }

            tmpf = Guid.NewGuid().ToString();

            string output = curTemplate.Replace("%%ANCHOR%%", OutputText.Text);

            File.WriteAllText(tmpfp + tmpf + ".tex", output);

            if (!File.Exists(tmpfp + tmpf + ".tex"))
            {
                MessageBox.Show("tex-file could not be created");
                return;
            }

            string[] latexArg = { 
                                    tmpf + ".tex",
                                    "-interaction=nonstopmode",
                                };

            runCommand(Properties.Settings.Default.PDFLatexPath, string.Join(" ", latexArg), tmpfp);

            if (!File.Exists(tmpfp + tmpf + ".pdf"))
            {
                MessageBox.Show("tex-file could not be parsed");
                return;
            }

            string[] gsArg = { 
                                 "-sDEVICE=pswrite",
                                 "-dNOCACHE",
                                 "-sOutputFile=" + tmpf + ".ps",
                                 "-q",
                                 "-dbatch",
                                 "-dNOPAUSE",
                                 tmpf + ".pdf",
                                 "-c",
                                 "quit"
                             };

            runCommand(Properties.Settings.Default.GSPath, string.Join(" ", gsArg), tmpfp);

            if (!File.Exists(tmpfp + tmpf + ".ps"))
            {
                MessageBox.Show("pdf-file could not be interpreted");
                return;
            }

            if (CDWin != null)
            {
                locked = true;

                StructImportOptions impOpts= new StructImportOptions();
                impOpts.MaintainLayers = true;
                impOpts.Mode = CorelDRAW.cdrImportMode.cdrImportFull;

                CDWin.ActiveLayer.ImportEx(tmpfp + tmpf + ".ps", CorelDRAW.cdrFilter.cdrPSInterpreted, impOpts).Finish();

                CorelDRAW.Shape newShape = CDWin.ActiveShape;

                newShape.Properties["latex2cd", (int)ShapeProperties.isLatexObject] = true;
                newShape.Properties["latex2cd", (int)ShapeProperties.latexText] = OutputText.Text;
                newShape.Properties["latex2cd", (int)ShapeProperties.latexTemplate] = Properties.Settings.Default.LastTemplate;

                OutputText.Text = "";

                if (oldShape != null)
                {
                    double d11, d12, d21, d22, tx, ty;
                    oldShape.GetMatrix(out d11, out d12, out d21, out d22, out tx, out ty);
                    newShape.SetMatrix(d11, d12, d21, d22, tx, ty);
                    newShape.OrderBackOf(oldShape);
                    newShape.Fill = oldShape.Fill;
                    newShape.Outline.CopyAssign(oldShape.Outline.GetCopy());
                    newShape.FillMode = oldShape.FillMode;
                    oldShape.Delete();
                }
                else
                {
                    newShape.AlignToPoint(CorelDRAW.cdrAlignType.cdrAlignVCenter, CDWin.Application.ActiveWindow.ActiveView.OriginX, CDWin.Application.ActiveWindow.ActiveView.OriginY);
                    newShape.AlignToPoint(CorelDRAW.cdrAlignType.cdrAlignHCenter, CDWin.Application.ActiveWindow.ActiveView.OriginX, CDWin.Application.ActiveWindow.ActiveView.OriginY);
                }

                locked = false;

                newShape.Selected = false;
                newShape.Selected = true;
            }
        }

        private void cleanTempFiles()
        {
            foreach (string f in Directory.GetFiles(tmpfp, tmpf + ".*", SearchOption.TopDirectoryOnly))
            {
                File.Delete(f);
            }
        }

        private void runCommand(string command, string Arguments, string WorkingDir)
        {
            Process prc = null;
            ProcessStartInfo psi = new ProcessStartInfo(command, Arguments);

            psi.WorkingDirectory = WorkingDir;
            psi.RedirectStandardInput = true;
            psi.RedirectStandardOutput = true;
            psi.WindowStyle = ProcessWindowStyle.Hidden;
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            psi.RedirectStandardError = true;

            prc = Process.Start(psi);

            prc.WaitForExit(100000);

            prc.Close();
        }

        private void TemplateButton_Click(object sender, RoutedEventArgs e)
        {
            TemplateWindow tmp = new TemplateWindow();
            tmp.ShowDialog();
        }

        public static bool ExistsOnPath(string fileName)
        {
            if (GetFullPath(fileName) != null)
                return true;
            return false;
        }

        public static string GetFullPath(string fileName)
        {
            if (File.Exists(fileName))
                return System.IO.Path.GetFullPath(fileName);

            var values = Environment.GetEnvironmentVariable("PATH");
            foreach (var path in values.Split(';'))
            {
                var fullPath = System.IO.Path.Combine(path, fileName);
                if (File.Exists(fullPath))
                    return fullPath;
            }
            return null;
        }
    }
}
