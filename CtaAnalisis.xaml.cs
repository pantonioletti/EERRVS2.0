/*
 * Created by SharpDevelop.
 * User: pantonio
 * Date: 02/16/2017
 * Time: 12:25
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using System.IO;
using NPOI.XSSF.UserModel;


namespace EstadoResultadoWPF
{
	/// <summary>
	/// Interaction logic for CtaAnalisis.xaml
	/// </summary>
	public partial class CtaAnalisis : Window
	{
		public CtaAnalisis()
		{
			InitializeComponent();
		}
		void btnPathIn_Click(object sender, RoutedEventArgs e)
		{
	 		FolderBrowserDialog fldIn = new FolderBrowserDialog();
            //fldIn.SelectedPath = "C:\\dev\\projects";
            fldIn.ShowDialog();
            //PathIn.Clear();
            ListInputFiles.Items.Clear();
            PathIn.Text = fldIn.SelectedPath;
            listFiles();
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
                    if (fName.EndsWith(".xls"))
                        //ListInputFiles.Items.Add(fName.Replace((PathIn.Text).Insert((PathOut.Text).Length, "\\"), ""));
                        ListInputFiles.Items.Add(fName.Replace(PathIn.Text+"\\", ""));
                }
            }
            catch (Exception excep)
            {
                System.Windows.Forms.MessageBox.Show(excep.Message);
            }

        }

        void btnProc_Click(object sender, RoutedEventArgs e){}
        
        void btnPathOut_Click(object sender, RoutedEventArgs e){}
	}
}