using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using ProcAnalisis;


namespace EstadoResultadoWPF
{
    public class AnalisisSS
    {
        public AnalisisSS(List<String> files, String outp)
        {
            XSSFWorkbook outWb = new XSSFWorkbook();
            ISheet outSheet = outWb.CreateSheet("Análisis");
            IRow r0 = outSheet.CreateRow(0);
            for (int i = 0; i < PAConstants.header.Length; i++)
            {
                ICell cell = r0.CreateCell(i);
                cell.SetCellValue(PAConstants.header[i]);
            }
            int rowNum = 1;
            foreach (String f in files)
            {
                IWorkbook wb = WorkbookFactory.Create(f);
                ISheet sheet = wb.GetSheetAt(0);
                IEnumerator re = sheet.GetRowEnumerator();
                re.MoveNext();
                HSSFRow r = (HSSFRow)re.Current;
                String cie_name = r.GetCell(1).StringCellValue;
                String cuenta = "";
                String c0 = "";
                while (re.MoveNext())
                {
                    r = (HSSFRow)re.Current;
                    try
                    {
                        c0 = r.GetCell(0).StringCellValue;
                    }
                    catch (NullReferenceException)
                    {
                        continue;
                    }
                    if (c0.Trim() == "")
                        continue;
                    if (c0.StartsWith("1."))
                        break;
                }
                cuenta = c0;
                while (re.MoveNext())
                {
                    r = (HSSFRow)re.Current;
                    try
                    {
                        c0 = r.GetCell(0).StringCellValue;
                    }
                    catch (NullReferenceException)
                    {
                        continue;
                    }

                    if (c0.Trim() == "")
                        continue;
                    if (c0.StartsWith("1."))
                        cuenta = c0;
                    else
                    {
                        int colNum = 0;
                        r0 = outSheet.CreateRow(rowNum++);
                        r0.CreateCell(colNum++).SetCellValue(cie_name);
                        r0.CreateCell(colNum++).SetCellValue(cuenta);
                        for (int i=0; i < 13; i++)
                        {
                            ICell c = r.GetCell(i);
                            if (c == null)
                                continue;
                            ICell nc = r0.CreateCell(colNum++);
                            nc.SetCellType(c.CellType);
                            if (c.CellType == NPOI.SS.UserModel.CellType.String)
                                nc.SetCellValue(c.StringCellValue);
                            else if (c.CellType == NPOI.SS.UserModel.CellType.Numeric)
                                nc.SetCellValue(c.NumericCellValue);
                            else
                                nc.SetCellValue(c.StringCellValue);
                        }

                    }
                    
                }
                wb.Close();
            }
            outWb.Write(new FileStream(outp, FileMode.Create, FileAccess.Write));
            outWb.Close();
        }
        /*public static void Main(String[] args)
        {
            List<String> lf = new List<String>();
            lf.Add("D:\\dev\\projects\\EERR\\data\\RCL ANALISIS 07.2019.xls");
            AnalisisSS a = new AnalisisSS(lf, "D:/dev/projects/EERR/data/pmas.xlsx");
        }*/
    }
}