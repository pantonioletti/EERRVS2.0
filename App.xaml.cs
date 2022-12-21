using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace EstadoResultadoWPF
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        [STAThread]
        public static void Main(String[] args)
        {
            //List<String> lf = new List<String>();
            //lf.Add("D:\\dev\\projects\\EERR\\data\\RCL ANALISIS 07.2019.xls");
            App app = new App();
            Console.WriteLine("Iniciando la aplicación");
            WEstadoResultado a = new WEstadoResultado();// lf, "D:/dev/projects/EERR/data/pmas.xlsx");
            a.Activate();
            a.Show();
            app.Run();
        }
    }

}
