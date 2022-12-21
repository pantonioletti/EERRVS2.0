using System;
using System.Linq;
using System.Text;
using System.IO;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PMAS;


/*
 * Created by SharpDevelop.
 * User: pantonio
 * Date: 21-11-2017
 * Time: 12:33
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */


 
 
class SaldoClientesRW : PMASLib
    {
/* Datos origen
 *  1- Cuenta: 1.1.20.1001 Deudores por Ventas Exportacio
 *  1- Cliente:	2016636
 *  2- Nombre Cliente:	 - 3 / 0 BLANCA ALVARADO DE ARANDIA
 *  3- Suc.	: 60
 *  4- Gru	: Tra
 *  5- Número: 42338
 *  6- Fecha : 17/02/2017	
 *  7- Vencto.: 30/05/2017
 *  8- Saldo por Vencer: 0
 *  9- Saldo Vencido: 7600
 * 10- Días: 123
 * 11- Saldo Documento: 7600
 * 
 */        
		const int C_IN_CUENTA = 0;
        const int C_IN_CLIENTE = 0;
        const int C_IN_NOMBRE = 1;
        const int C_IN_SUC = 2;
        const int C_IN_GRU = 3;
        const int C_IN_NUM = 4;
        const int C_IN_FECHA = 5;
        const int C_IN_VCTO = 6;
        const int C_IN_SALDOPOR = 7;
        const int C_IN_SALDOV = 8;
        const int C_IN_DIAS = 9;
        const int C_IN_SALDODOC = 10;

/* Datos destino:
 *  1- Cuenta (1) Deudores por Venta
 *  2- Número (1) 1.1.20.1001
 *  3- Cuenta (1) Deudores por Ventas Exportacio
*  4- Cliente (1) 2016636
 *  5- Suc.   (2) 3 / 0
 *  6- Nombre o Razón Social (2) BLANCA ALVARADO DE ARANDIA
 *  7- Sucursal (3): 60
 *  8- Grupo	(4): Tra
 *  9- Número (5): 42228
 * 10- Fecha (6) 17/02/2017
 * 11- Vencto. (7) 30/05/2017
 * 12- Saldo por Vencer (8) 0
 * 13- Saldo Vencido (9) 7600
 * 14- Días (10) 123
 * 15- Saldo Documento (11) 7600
 * 16- Dato - <Fecha arbitraria>-(7)
* 17- Periodo- ??
  */
        const int C_OUT_CTAS = 0;
        const int C_OUT_NUM = 1;
        const int C_OUT_CTAL = 2;
        const int C_OUT_CLTE = 3;
        const int C_OUT_SUC = 4;
        const int C_OUT_NOMBRE = 5;
        const int C_OUT_SUCURSAL = 6;
        const int C_OUT_GRUPO = 7;
        const int C_OUT_NUMERO = 8;
        const int C_OUT_FECHA = 9;
        const int C_OUT_VCTO = 10;
        const int C_OUT_SALDOP = 11;
        const int C_OUT_SALDOV = 12;
        const int C_OUT_DIAS = 13;
        const int C_OUT_SALDOD = 14;
        const int C_OUT_DATO = 15;
        const int C_OUT_PERIODO = 16;

        //string[] months ={ "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE" };

        const string C_STR_IN_HEAD = "RCL SUDAMERICANA SOCIEDAD ANONIMA";
        //const string C_STR_IN_ACCOUNT = "Cuenta Contable";
        //const char C_COL_SEPARATOR = ';';

        //const string C_ERR_MSG_FILE_FMT_ERR = "File format incorrect";

        //const string C_DATA_STATUS = "REAL";
        
        
        public StringBuilder readXlsx(string file, /*EERRDataAndMethods eerr,*/ XSSFWorkbook twb)
        {
            StringBuilder retVal = new StringBuilder("");
            IWorkbook wb = null;
            
            FileInfo fi = new FileInfo(file);
            if (file.EndsWith(".xlsx"))
            	wb = new XSSFWorkbook(fi);
            else if (file.EndsWith(".xls"))
            	wb = new HSSFWorkbook(new FileStream(file, FileMode.Open));

            ISheet sheet = wb.GetSheetAt(0);
            int rowNum = 0;
            IRow r = sheet.GetRow(rowNum);
            ICell c = r.GetCell(0);
            string company = c.StringCellValue;
            XSSFSheet sh = (XSSFSheet)twb.GetSheet("Data");

            while(++rowNum < sheet.LastRowNum)
            {
		        r = sheet.GetRow(rowNum);
	            c = r.GetCell(0);
	            if (c.StringCellValue.Trim().Equals("Cliente"))
	                break;
            	
            }
            while (++rowNum < sheet.LastRowNum)
            {
            	String sCuenta="";
            	String sNum="";
            	String sCuenta2="";
            	String sCliente="";
            	String sSuc = "";
            	String sNombre = "";
            	

	            r = sheet.GetRow(rowNum);
	            c = r.GetCell(C_IN_CUENTA);
	            String value = getCellValue(c).Trim();
	            if (value.Equals("")){
	            	String aux = value;
	            	c = r.GetCell(C_IN_SUC);
	            	value = getCellValue(c).Trim();
	            	if (value.Equals(""))
	                	continue;
	            	value = aux;
	            }
	            int iCuenta;
            	if (int.TryParse(value, out iCuenta)) //Nro de cliente
            	{
                    sCliente = value;
                    
            		sSuc="0";
            		sNombre = "NN";

            		c=r.GetCell(C_IN_NOMBRE); // Nombre del cliente
            		value = getCellValue(c); // "- 3 / 0 BLANCA ALVARADO DE ARANDIA"
            		value = value.Remove(0,3); // "3 / 0 BLANCA ALVARADO DE ARANDIA"
            		int pos;
            		for(pos = value.IndexOf('/')+2;pos < value.Length && value[pos] != ' ';pos++);
            		if (pos < value.Length){
                        sSuc = value.Substring(0,pos); //"3 / 0"
                        sNombre = value.Substring(pos+1); // "BLANCA ALVARADO DE ARANDIA"
            		}
            	}else if (value.Trim().Length > 0){ //Numero de cuenta
            		String[] ss = value.Split(' ');
            		sCuenta=ss[0].Substring(0,18);
            		sCuenta2=ss[0];
            		sNum = ss[1];
            		continue;
            	}

                XSSFRow row = (XSSFRow)sh.CreateRow(sh.LastRowNum+1);

                XSSFCell cell = (XSSFCell)row.CreateCell(C_OUT_CTAS);
                cell.SetCellValue(sCuenta);
                cell = (XSSFCell)row.CreateCell(C_OUT_NUM);
                cell.SetCellValue(sNum);
                cell = (XSSFCell)row.CreateCell(C_OUT_CTAL);
                cell.SetCellValue(sCuenta2);
                
                cell = (XSSFCell)row.CreateCell(C_OUT_CLTE);
                cell.SetCellValue(sCliente);
                cell.SetCellType(CellType.Numeric);
                
                cell = (XSSFCell)row.CreateCell(C_OUT_SUC);
                cell.SetCellValue(sSuc);//"3 / 0"
                cell = (XSSFCell)row.CreateCell(C_OUT_NOMBRE);
                cell.SetCellValue(sNombre);// "BLANCA ALVARADO DE ARANDIA"
            		
        		c=r.GetCell(C_IN_SUC); // Sucursal
                cell = (XSSFCell)row.CreateCell(C_OUT_SUCURSAL);
                cell.SetCellValue(getCellValue(c));
        		
        		c=r.GetCell(C_IN_GRU); // Grupo
                cell = (XSSFCell)row.CreateCell(C_OUT_GRUPO);
                cell.SetCellValue(getCellValue(c));

        		c=r.GetCell(C_IN_NUM); // Numero
                cell = (XSSFCell)row.CreateCell(C_OUT_NUMERO);
                cell.SetCellValue(getCellValue(c));
        		
        		c=r.GetCell(C_IN_FECHA); // Fecha
                cell = (XSSFCell)row.CreateCell(C_OUT_FECHA);
                cell.SetCellValue(getCellDateValue(c));

        		c=r.GetCell(C_IN_VCTO); // Vencimiento
                cell = (XSSFCell)row.CreateCell(C_OUT_VCTO);
                cell.SetCellValue(getCellDateValue(c));
                
        		c=r.GetCell(C_IN_SALDOPOR); // Saldo por vencer
                cell = (XSSFCell)row.CreateCell(C_OUT_SALDOP);
                cell.SetCellValue(getCellValue(c));
                cell.SetCellType(CellType.Numeric);
                
        		c=r.GetCell(C_IN_SALDOV); // Saldo vencido
                cell = (XSSFCell)row.CreateCell(C_OUT_SALDOV);
                cell.SetCellValue(getCellValue(c));
                cell.SetCellType(CellType.Numeric);

        		c=r.GetCell(C_IN_DIAS); // Dias
                cell = (XSSFCell)row.CreateCell(C_OUT_DIAS);
                cell.SetCellValue(getCellValue(c));
                cell.SetCellType(CellType.Numeric);
                
        		c=r.GetCell(C_IN_SALDODOC); // Dias
                cell = (XSSFCell)row.CreateCell(C_OUT_SALDOD);
                cell.SetCellValue(getCellValue(c));
                cell.SetCellType(CellType.Numeric);

            	}
            

            return retVal;
        	
        }
        
        
    }