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

namespace EstadoResultadoRCL
{
    /// <summary>
    /// Interaction logic for ParamMaint.xaml
    /// </summary>
    public partial class ParamMaint : Window
    {
        private string curParam;
        private EERRDataAndMethods eerr;
        public ParamMaint(string param, EERRDataAndMethods eerr)
        {
            InitializeComponent();
            curParam = param;
            this.eerr = eerr;
            initialize();
        }
        private void initialize()
        {
            if (curParam.Equals(Constants.INV_ITEMS))
            {

                dgParamsData.Width = 360;
                this.Width = 390;

                List<ItemData> lItem = eerr.getItems();
                dgParamsData.ItemsSource = lItem;



            }
            else if (curParam.Equals(Constants.INV_AREAS))
            {
                dgParamsData.Width = 360;
                this.Width = 390;

                List<AreaData> lArea = eerr.getAreas();
                dgParamsData.ItemsSource = lArea;

            }
            else if (curParam.Equals(Constants.INV_LINEAS))
            {

            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (curParam.Equals(Constants.INV_AREAS))
            {
                dgParamsData.Items.MoveCurrentToFirst();

                while (!dgParamsData.Items.IsCurrentAfterLast)
                {
                    AreaData area = (AreaData)dgParamsData.Items.CurrentItem;
                    dgParamsData.Items.MoveCurrentToNext();
                    string[] recArea = eerr.getArea(area.Area);
                    if (recArea == null || !area.Marca.Equals(recArea[0]) || !area.Agrupacion.Equals(recArea[1]))
                        eerr.setArea(area);
                }
            }
            else if (curParam.Equals(Constants.INV_ITEMS))
            {
                dgParamsData.Items.MoveCurrentToFirst();
                List<ItemData> id = new List<ItemData>();
                while (!dgParamsData.Items.IsCurrentAfterLast)
                {
                    ItemData item = (ItemData)dgParamsData.Items.CurrentItem;
                    id.Add(item);
                    dgParamsData.Items.MoveCurrentToNext();

                }
                eerr.updateItems(id);
            }
            else if (curParam.Equals(Constants.INV_LINEAS))
            {

            }
        }
    }
}
