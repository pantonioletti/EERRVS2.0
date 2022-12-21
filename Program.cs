using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Packaging;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Diagnostics;

namespace ReadZip
{
    public class XlReader
    {
        protected SpreadsheetDocument xlDoc;
        protected SharedStringItem[] sst;
        protected struct WSheet { public string name; public WorksheetPart ws;};
        protected WSheet[] mySheets;
        protected CellFormat[] myFormats;



        ~XlReader()
        {
            xlDoc.Close();
        }


        public XlReader(string xlFile)
        {
            xlDoc = SpreadsheetDocument.Open(xlFile, false);
            IEnumerator<WorksheetPart> wsEnum = xlDoc.WorkbookPart.WorksheetParts.GetEnumerator();
            WorksheetPart[] wsp = xlDoc.WorkbookPart.WorksheetParts.ToArray();
            mySheets = new WSheet[wsp.Length];

            int id = 0;
            foreach (OpenXmlElement el in xlDoc.WorkbookPart.Workbook.Sheets)
            {
                foreach (OpenXmlAttribute attr in el.GetAttributes())
                {
                    if (attr.LocalName.Equals("name"))
                    {
                        mySheets[id].name = attr.Value;
                        mySheets[id].ws = wsp[id];
                        break;
                    }
                }
                id++;
            }
            sst = xlDoc.WorkbookPart.SharedStringTablePart
                            .SharedStringTable.Elements<SharedStringItem>()
                            .ToArray<SharedStringItem>();
            CellFormats cfmts = xlDoc.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats;
            myFormats = new CellFormat[cfmts.Count];
            
            id=0;
            foreach(OpenXmlElement stEl in cfmts.ChildElements)
                myFormats[id++] = new CellFormat(stEl.OuterXml);

        }

    }
}
