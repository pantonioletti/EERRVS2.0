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
using System.Windows.Forms;

using NPOI.XSSF.UserModel;
using System.IO;

namespace EstadoResultadoRCL
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private EERRDataAndMethods eerrLib;
        public MainWindow()
        {
            InitializeComponent();
            string[] cmdLn = Environment.GetCommandLineArgs();
            if (cmdLn.Length < 2)
                System.Windows.Forms.MessageBox.Show("El aplicativo require archivo de configuración (.ini).");
            else
                eerrLib = new EERRDataAndMethods(cmdLn[1]);
            string s = eerrLib.getIniParam(Constants.DEFAULT_INPUT_DIR);
            if (string.IsNullOrEmpty(s))
                s = "";
            PathIn.Text = s;

            s = eerrLib.getIniParam(Constants.DEFAULT_OUTPUT_DIR);
            if (string.IsNullOrEmpty(s))
                s = "";
            PathOut.Text = s;
            listFiles();
        }

        private void mnItem_Click(object sender, RoutedEventArgs e)
            {
                ParamMaint itemMaint = new ParamMaint(Constants.INV_ITEMS, eerrLib);
                itemMaint.Title = "Mantención de Items";
                itemMaint.Show();
            }

            private void mnAreas_Click(object sender, RoutedEventArgs e)
            {
                ParamMaint itemMaint = new ParamMaint(Constants.INV_AREAS, eerrLib);
                itemMaint.Title = "Mantención de Areas";
                itemMaint.Show();
            }

            private void listFiles()
            {
                try
                {
                    ListInputFiles.Items.Clear();
                    IEnumerable<string> files = Directory.EnumerateFiles(PathIn.Text);//, "*.csv,*.xls");
                    IEnumerator<string> enFiles = files.GetEnumerator();
                    while (enFiles.MoveNext())
                    {
                        string fName = enFiles.Current;
                        //if (fName.EndsWith(".csv") || fName.EndsWith(".xls"))
                        if (fName.EndsWith(".xls") || fName.EndsWith(".xlsx"))
                            //ListInputFiles.Items.Add(fName.Replace((PathIn.Text).Insert((PathOut.Text).Length, "\\"), ""));
                            ListInputFiles.Items.Add(fName.Replace(PathIn.Text + "\\", ""));
                    }
                }
                catch (Exception excep)
                {
                    System.Windows.Forms.MessageBox.Show(excep.Message);
                }

            }
            private void btnPathIn_Click(object sender, RoutedEventArgs e)
            {
                FolderBrowserDialog fldIn = new FolderBrowserDialog();
                //fldIn.SelectedPath = "C:\\dev\\projects";
                fldIn.ShowDialog();
                //PathIn.Clear();
                ListInputFiles.Items.Clear();
                PathIn.Text = fldIn.SelectedPath;
                listFiles();
            }
            private void btnPathOut_Click(object sender, RoutedEventArgs e)
            {
                FolderBrowserDialog fldOut = new FolderBrowserDialog();
                //fldOut.SelectedPath = "C:\\dev\\projects";
                fldOut.ShowDialog();
                PathOut.Text = fldOut.SelectedPath;
            }

            private void btnProc_Click(object sender, RoutedEventArgs e)
            {
                if (PathIn.Text.Length == 0 || ListInputFiles.SelectedItems.Count == 0)
                    System.Windows.MessageBox.Show("Debe haber archivos de entrada seleccionados");
                else if (PathOut.Text.Length == 0)
                    System.Windows.MessageBox.Show("Debe seleccionar carpeta e salida");
                else if (chkApplyRate.IsChecked.Value && txtExchgRate.Text.Trim().Length == 0)
                {
                    System.Windows.MessageBox.Show("Debe ingresar el tipo de cambio");
                    txtExchgRate.Focus();
                }
                else
                {
                    FileOut.Clear();
                    StringBuilder s = new StringBuilder();
                    s.Append("eerr_");
                    s.Append(System.DateTime.Now.Ticks.ToString());
                    FileOut.Text = s.ToString() + ".xlsx";

                    StreamWriter log = new StreamWriter(s.ToString() + ".log");
                    XSSFWorkbook xlDoc = new XSSFWorkbook();

                    XSSFSheet sh = (XSSFSheet)xlDoc.CreateSheet(Constants.EERR_SHEET_NAME);
                    var r = sh.CreateRow(0);
                    int col = 0;
                    foreach (string h in Constants.EERR_SHEET_HEADERS)
                    {
                        (r.CreateCell(col++)).SetCellValue(h);
                    }

                    System.Collections.IEnumerator idxEnum = ListInputFiles.SelectedItems.GetEnumerator();
                    EERRCsvRW csvrw = new EERRCsvRW();
                    string path = PathIn.Text + "\\";
                    while (idxEnum.MoveNext())
                    {
                        string inFile = path + (string)idxEnum.Current;
                        Double rate = chkApplyRate.IsChecked.Value ? Double.Parse(txtExchgRate.Text) : 0;
                        csvrw.readXlsx(inFile, eerrLib, xlDoc, chkApplyRate.IsChecked.Value, rate);
                    }
                    xlDoc.Write(new FileStream(PathOut.Text + "\\" + FileOut.Text /*+ ".xlsx"*/, FileMode.Create, FileAccess.Write));
                    log.Close();
                    System.Windows.MessageBox.Show("Proceso terminado");
                }
            }

            private void chkApplyRate_Checked(object sender, RoutedEventArgs e)
            {

            }
            private void chkApplyRate_Click(object sender, RoutedEventArgs e)
            {
                if (chkApplyRate.IsChecked.Value)
                {
                    txtExchgRate.Text = "";
                    txtExchgRate.IsEnabled = true;
                    txtExchgRate.Focus();
                }
                else
                {
                    txtExchgRate.IsEnabled = false;
                }
            }

            private void txtExchgRate_TextChanged(object sender, TextChangedEventArgs e)
            {
                if (txtExchgRate.Text != "")
                {
                    try
                    {
                        double rate = Double.Parse(txtExchgRate.Text);
                    }
                    catch (FormatException fe)
                    { txtExchgRate.Text = "Un número CTM!!!!!"; }
                    catch (OverflowException oe)
                    { txtExchgRate.Text = "No tan grande saco 'e coco!!!!"; }
                }

            }
            void mnConfig_Click(object sender, RoutedEventArgs e)
            {
                throw new NotImplementedException();
            }

            void mnCtaAnalisis_Click(object sender, RoutedEventArgs e)
            {
                CtaAnalisis ctaAnalisis = new CtaAnalisis();
                ctaAnalisis.Title = "Mantención de Items";
                ctaAnalisis.Show();
            }

            void mnCtasCltes_Click(object sender, RoutedEventArgs e)
            {

            }

        }

    }
