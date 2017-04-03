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
        bool _latexBin = false;
        bool _gsBin = false;
        bool latexBin
        {
            get { return _latexBin; }
            set
            {
                checkPDF.IsChecked = value;
                _latexBin = value;
            }
        }
        bool gsBin
        {
            get { return _gsBin; }
            set
            {
                checkGS.IsChecked = value;
                _gsBin = value;
            }
        }
        bool hideWin
        {
            get { return Properties.Settings.Default.hideWindows; }
            set
            {
                checkHide.IsChecked = value;
                Properties.Settings.Default.hideWindows = value;
                Properties.Settings.Default.Save();
            }
        }
        bool showLog
        {
            get { return Properties.Settings.Default.showLog; }
            set
            {
                checkLog.IsChecked = value;
                Properties.Settings.Default.showLog = value;
                Properties.Settings.Default.Save();
            }
        }



        public LatexDock()
        {
            InitializeComponent();
            initDock();
        }
        public LatexDock(object app)
        {
            InitializeComponent();
            initDock();

            CDWin = (CorelDRAW.Application)app;
            CDWin.SelectionChange += CDWin_SelectionChange;
            getSelection();
        }

        private void initDock()
        {
            checkPath();
            getTemplate();
            checkHide.IsChecked = hideWin;
            checkLog.IsChecked = showLog;
        }


        private void checkPath()
        {
            latexBin = ExistsOnPath(Properties.Settings.Default.PDFLatexPath);
            gsBin = ExistsOnPath(Properties.Settings.Default.GSPath);
        }

        private void getTemplate()
        {
            if (!Properties.Settings.Default.LastTemplate.Contains("%%ANCHOR%%"))
                curTemplate = Properties.Settings.Default.DefaultTemplate;
            else
                curTemplate = Properties.Settings.Default.LastTemplate;
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

            string LatexLog = "";
            runCommand(Properties.Settings.Default.PDFLatexPath, string.Join(" ", latexArg), tmpfp, out LatexLog);

            if (!File.Exists(tmpfp + tmpf + ".pdf"))
            {
                MessageBox.Show("tex-file could not be parsed. Press OK to see Log.");
                if (LatexLog != "")
                    displayLog(LatexLog);
                return;
            }

            string[] gsArg = { 
                                 "-sDEVICE=ps2write",
                                 "-dNOCACHE",
                                 "-sOutputFile=" + tmpf + ".ps",
                                 "-q",
                                 "-dbatch",
                                 "-dNOPAUSE",
                                 tmpf + ".pdf",
                                 "-c",
                                 "quit",
                             };

            string GSout = "";
            runCommand(Properties.Settings.Default.GSPath, string.Join(" ", gsArg), tmpfp, out GSout);

            if (!File.Exists(tmpfp + tmpf + ".ps"))
            {
                MessageBox.Show("pdf-file could not be interpreted");
                if (GSout != "")
                    displayLog(LatexLog + Environment.NewLine + GSout);
                return;
            }

            if (showLog)
            {
                displayLog(LatexLog + Environment.NewLine + GSout);
            }

            if (CDWin != null)
            {
                if (CDWin.ActiveLayer == null) return;
                locked = true;

                StructImportOptions impOpts = new StructImportOptions();
                impOpts.MaintainLayers = true;
                impOpts.Mode = CorelDRAW.cdrImportMode.cdrImportFull;

                CDWin.ActiveLayer.ImportEx(tmpfp + tmpf + ".ps", CorelDRAW.cdrFilter.cdrPSInterpreted, impOpts).Finish();

                CorelDRAW.Shape newShape = CDWin.ActiveShape;

                newShape.Properties["latex2cd", (int)ShapeProperties.isLatexObject] = true;
                newShape.Properties["latex2cd", (int)ShapeProperties.latexText] = OutputText.Text;
                newShape.Properties["latex2cd", (int)ShapeProperties.latexTemplate] = curTemplate;

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

        private void displayLog(string log)
        {
            if (log.Trim() == "") return;
            new errorLog(log);
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
            string trash;
            runCommand(command, Arguments, WorkingDir, out trash);
        }

        private void runCommand(string command, string Arguments, string WorkingDir, out string output)
        {
            Process prc = null;
            ProcessStartInfo psi = new ProcessStartInfo(command, Arguments);

            psi.WorkingDirectory = WorkingDir;

            psi.RedirectStandardInput = true;
            psi.RedirectStandardOutput = true;
            psi.WindowStyle = (hideWin) ? ProcessWindowStyle.Hidden : ProcessWindowStyle.Normal;
            psi.UseShellExecute = false;
            psi.CreateNoWindow = hideWin;
            psi.RedirectStandardError = true;

            prc = Process.Start(psi);

            prc.WaitForExit(100000);

            output = command + " " + Arguments + Environment.NewLine + prc.StandardOutput.ReadToEnd() + Environment.NewLine;

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

        private void pdfBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            dlg.DefaultExt = ".exe";
            dlg.Filter = "pdflatex.exe|pdflatex.exe";
            bool? result = dlg.ShowDialog();

            if (result == true && ExistsOnPath(dlg.FileName))
            {
                Properties.Settings.Default.PDFLatexPath = dlg.FileName;
                Properties.Settings.Default.Save();
            }
            checkPath();
        }

        private void gsBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            dlg.DefaultExt = ".exe";
            dlg.Filter = "Ghostscript|gs.exe;mgs.exe";
            bool? result = dlg.ShowDialog();

            if (result == true && ExistsOnPath(dlg.FileName))
            {
                Properties.Settings.Default.GSPath = dlg.FileName;
                Properties.Settings.Default.Save();
            }
            checkPath();
        }

        private void checkHide_Click(object sender, RoutedEventArgs e)
        {
            hideWin = !hideWin;
        }

        private void checkLog_Click(object sender, RoutedEventArgs e)
        {
            showLog = !showLog;
        }
    }
}
