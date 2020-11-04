using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using System.Data.SQLite;
using System.Collections;

namespace EstadoResultadoRCL
{
    public class ConfigModel
    {
        private static string QUERY_ITEMS = "select cod, desc from items;";
        private static string QUERY_AREA = "select area, marca, agrupacion from area;";
        private static string QUERY_EERR = "select length(prefix) l, prefix, desc from eerr order by l asc, prefix asc;";
        private static string QUERY_SUCURSAL = "select cod, desc from sucursal;";

        private SQLiteDataAdapter ad;
        private SQLiteConnection sqlite = null;
        private string dbFile;

        private SQLiteConnection getConnection()
        {
            SQLiteConnection retVal = null;
            if (sqlite == null || sqlite.State == ConnectionState.Broken || sqlite.State == ConnectionState.Closed)
                sqlite = new SQLiteConnection("Data Source=" + dbFile);
            if (sqlite.State != ConnectionState.Open)
                sqlite.Open();

            retVal = sqlite;
            return retVal;

        }

        public Dictionary<string, string> getBranchs()
        {
            Dictionary<string, string> branchs = new Dictionary<string, string>();
            SQLiteConnection conn = getConnection();
            System.Data.DataTable dt = new System.Data.DataTable();
            try
            {
                SQLiteCommand cmd;
                //conn.Open();  //Initiate connection to the db

                cmd = conn.CreateCommand();
                cmd.CommandText = QUERY_SUCURSAL;
                ad = new SQLiteDataAdapter(cmd);
                ad.Fill(dt); //fill the datasource
                DataRow[] rows = dt.Select();
                for (int i = 0; i < rows.Length; i++)
                {
                    branchs.Add((string)rows[i][Constants.BRANCH_1], (string)rows[i][Constants.BRANCH_2]);
                }
                ad.Dispose();
                dt.Dispose();
                rows = null;

            }
            catch (SQLiteException ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            sqlite = null;
            return branchs;

        }


        public Dictionary<string, string> getItems()
        {
            Dictionary<string, string> items = new Dictionary<string, string>();
            SQLiteConnection conn = getConnection();
            System.Data.DataTable dt = new System.Data.DataTable();
            try
            {
                SQLiteCommand cmd;
                //conn.Open();  //Initiate connection to the db

                cmd = conn.CreateCommand();
                cmd.CommandText = QUERY_ITEMS;
                ad = new SQLiteDataAdapter(cmd);
                ad.Fill(dt); //fill the datasource
                DataRow[] rows = dt.Select();
                for (int i = 0; i < rows.Length; i++)
                {
                    items.Add((string)rows[i][Constants.ITEMS_1], (string)rows[i][Constants.ITEMS_2]);
                }
                ad.Dispose();
                dt.Dispose();
                rows = null;

            }
            catch (SQLiteException ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            sqlite = null;
            return items;

        }

        public int execQueries(List<string> stmts)
        {
            int retVal = 0;
            SQLiteConnection conn = getConnection();
            try
            {
                SQLiteCommand cmd;
                cmd = conn.CreateCommand();
                foreach (string stmt in stmts)
                {
                    cmd.CommandText = stmt;
                    retVal += cmd.ExecuteNonQuery();
                }

            }
            catch (SQLiteException ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                retVal = -1;
            }
            return retVal;

        }

        public Dictionary<string, string[]> getAreas()
        {
            Dictionary<string, string[]> areas = new Dictionary<string, string[]>();
            SQLiteConnection conn = getConnection();
            System.Data.DataTable dt = new System.Data.DataTable();
            try
            {
                SQLiteCommand cmd;
                //conn.Open();  //Initiate connection to the db

                cmd = conn.CreateCommand();
                cmd.CommandText = QUERY_AREA;
                ad = new SQLiteDataAdapter(cmd);
                dt = new System.Data.DataTable();
                ad.Fill(dt); //fill the datasource
                DataRow[] rows = dt.Select();
                for (int i = 0; i < rows.Length; i++)
                {
                    string[] marca_agrup = new string[2];
                    marca_agrup[0] = (string)rows[i][Constants.AREA_2];
                    marca_agrup[1] = (string)rows[i][Constants.AREA_3];
                    areas.Add((string)rows[i][Constants.AREA_1], marca_agrup);
                }
                ad.Dispose();
                dt.Dispose();

            }
            catch (SQLiteException ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            sqlite = null;

            return areas;

        }

        public ArrayList getEERRs()
        {
            ArrayList eerrs = new ArrayList();
            SQLiteConnection conn = getConnection();
            System.Data.DataTable dt = new System.Data.DataTable();
            try
            {
                SQLiteCommand cmd;
                //conn.Open();  //Initiate connection to the db

                cmd = conn.CreateCommand();
                cmd.CommandText = QUERY_EERR;  //set the passed query
                ad = new SQLiteDataAdapter(cmd);
                dt = new System.Data.DataTable();
                ad.Fill(dt); //fill the datasource
                DataRow[] rows = dt.Select();
                for (int i = 0; i < rows.Length; i++)
                {
                    string[] prefix_desc = new string[2];
                    prefix_desc[0] = (string)rows[i][Constants.EERR_1];
                    prefix_desc[1] = (string)rows[i][Constants.EERR_2];
                    eerrs.Add(prefix_desc);
                }
                ad.Dispose();
                dt.Dispose();

            }
            catch (SQLiteException ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            sqlite = null;
            return eerrs;

        }


        public ConfigModel(string db_file)
        {
            dbFile = db_file;
            sqlite = getConnection();
            System.Data.DataTable dt = new System.Data.DataTable();

            /*try
            {

                // Fourth load branches
                cmd = sqlite.CreateCommand();
                cmd.CommandText = QUERY_SUCURSAL;  //set the passed query
                ad = new SQLiteDataAdapter(cmd);
                dt = new System.Data.DataTable();
                ad.Fill(dt); //fill the datasource
                rows = dt.Select();
                for (int i = 0; i < rows.Length; i++)
                {
                    sucursal.Add((string)rows[i][Constants.BRANCH_1], (string)rows[i][Constants.BRANCH_2]);
                }
                arrEERR = eerr.ToArray();
                ad.Dispose();
            }
            catch (SQLiteException ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            sqlite.Close();*/

        }
        ~ConfigModel()
        {
            if (sqlite != null && sqlite.State != ConnectionState.Broken && sqlite.State != ConnectionState.Closed)
                sqlite.Close();
        }
    }
}
