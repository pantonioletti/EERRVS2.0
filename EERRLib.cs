using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EstadoResultadoWPF
{
	public class AreaData
	{
		public string Area {get; set;}
		public string Marca {get; set;}
		public string Agrupacion {get; set;}
	}

	public class ItemData
	{
		public string Codigo {get; set;}
		public string Descripcion {get; set;}
	}
	
    public class EERRDataAndMethods
    {
        private SQLiteConnection sqlite;
        private Dictionary<string, string> items = new Dictionary<string, string>();
        private Dictionary<string, string[]> area = new Dictionary<string, string[]>();
        private Dictionary<string, string> sucursal = new Dictionary<string, string>();
        Dictionary<string, string> confKeyValuePairs = new Dictionary<string, string>();
        private Object[] arrEERR;
        ConfigModel cfgModel;

        public EERRDataAndMethods(string confFile)
        {
            loadConf(confFile);
        }

        private void loadConf(string confFile)
        {
            confKeyValuePairs.Clear();
            string db_file = "";
            // Load conf file
            try
            {
                StreamReader sr = new StreamReader(confFile);

                string line;
                while (!sr.EndOfStream)
                {
                    line = sr.ReadLine();
                    if (line.Contains("="))
                    {
                        string[] keyValue = line.Split('=');
                        confKeyValuePairs.Add(keyValue[0].Trim(), keyValue[1].Trim());
                    }
                }
                sr.Close();
                if (!confKeyValuePairs.TryGetValue(Constants.DBFILE, out db_file))
                {
                    System.Windows.Forms.MessageBox.Show("No se encontraron datos de EERR");
                    return;
                }
            }
            catch (FileNotFoundException ex)
            {

                System.Windows.Forms.MessageBox.Show("El aplicativo require archivo de configuración (.ini).");
                System.Console.WriteLine(ex.Message);
            }
            
            cfgModel = new ConfigModel(db_file);
            
            items = cfgModel.getItems();
            area = cfgModel.getAreas();
            arrEERR = cfgModel.getEERRs().ToArray();
            sucursal = cfgModel.getBranchs();
            
        }

        public string getIniParam(string key)
        {
            string retVal = null;

            if (!string.IsNullOrEmpty(key) && confKeyValuePairs.ContainsKey(key))
                retVal = confKeyValuePairs[key];
            return retVal;
        }
 
        public List<ItemData> getItems()
        {
        	List<ItemData> lItem = new List<ItemData>();
            foreach(string key in items.Keys)
            {
            	string vDesc = items[key];
            	lItem.Add(new ItemData {Codigo =key, Descripcion = vDesc});
            }
            return lItem;
        }
        
        public bool updateItems(List<ItemData> plItem)
        {
        	bool retVal = false;
        	List<string> stmts=new List<string>();
        	
        	foreach(ItemData id in plItem)
        	{
        		if (items.ContainsKey(id.Codigo)){
        		    if (!items[id.Codigo].Equals(id.Descripcion))
        		    	stmts.Add("update items set desc = '"+id.Descripcion+"';");
        		}
        		else
        			stmts.Add("insert into items (cod,desc) values ('"+id.Codigo+"','"+id.Descripcion+"';");
        	}
        	if (cfgModel.execQueries(stmts) >= 0)
        		retVal = true;
        	
        	return retVal;
        	
        }

       
        public List<AreaData> getAreas()
        {
        	List<AreaData> lArea = new List<AreaData>();
            foreach(string key in area.Keys)
            {
            	string[] vArea = area[key];
            	lArea.Add(new AreaData {Area =key, Marca = vArea[0], Agrupacion = vArea[1]});
            }
            return lArea;
        }

        public string[] getArea(string code)
        {
        	string[] retVal = null;
            if (area.ContainsKey(code))
            {
                retVal = area[code];
            }
            return retVal;
        }
        
        public string getMarca(string pArea)
        {
        	string retVal = null;
            if (area.ContainsKey(pArea))
            {
            	retVal = area[pArea][0];
            }
            return retVal;
        	
        }
        
        public string getAgrupacion(string pArea)
        {
        	string retVal = null;
            if (area.ContainsKey(pArea))
            {
            	retVal = area[pArea][1];
            }
            return retVal;
        	
        }

        public void setArea(AreaData ad)
        {
        	if (area.ContainsKey(ad.Area))
        	{
        		string[] ma = area[ad.Area];
        		ma[0] = ad.Marca;
        		ma[1] = ad.Agrupacion;
        	}
        	else{
        		this.area.Add(ad.Area, new string[]{ad.Marca, ad.Agrupacion});
        	}
        }

        public string getBrand(string code)
        {
            string retVal = "N/A";
            if (!string.IsNullOrEmpty(code) && area.ContainsKey(code))
            {
                retVal = area[code][0];
            }
            return retVal;
        }

        public Object[] getLineas()
        {
            return arrEERR;
        }

        public string[] getLinea(string acct)
        {
            string[] retVal = new string[3] {null, null, null};
            for (int i = 0; i < arrEERR.Length; i++)
                if (acct.StartsWith(((string[])arrEERR[i])[0]))
                {
                    retVal[0] = ((string[])arrEERR[i])[1];
                    retVal[1] = ((string[])arrEERR[i])[2];
                    retVal[2] = ((string[])arrEERR[i])[3];
                    break;
                }
            return retVal;
        }

        public string getItem(string code)
        {
            string retVal = "N/A";
            if (!string.IsNullOrEmpty(code) && items.ContainsKey(code))
            {
                retVal = items[code];
            }
            return retVal;
        }

        public string getSucursal(string cod)
        {
            string retVal = "N/A";
            if (sucursal.ContainsKey(cod))
            {
                retVal = sucursal[cod];
            }
            return retVal;
        }

        public SpreadsheetDocument buildSpreadsheet(string filename)
        {
            SpreadsheetDocument xlDoc;
            try
            {
                xlDoc = SpreadsheetDocument.Create(filename + ".xlsx", SpreadsheetDocumentType.Workbook);
                WorkbookPart wbp = xlDoc.AddWorkbookPart();
                wbp.Workbook = new Workbook();

                wbp.AddNewPart<WorksheetPart>();

                Sheets shts = wbp.Workbook.AppendChild<Sheets>(new Sheets());
                wbp.Workbook.Save();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                xlDoc = null;
            }
            return xlDoc;
        }


    }
}

