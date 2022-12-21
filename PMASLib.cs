/*
 * Created by SharpDevelop.
 * User: pantonio
 * Date: 11/28/2017
 * Time: 13:23
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace PMAS
{
	public abstract class PMASLib{
        string[] months ={ "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE" };
	    public String getCellValue(ICell c)
        {
        	String retVal = "";
        	if (c != null)
        		retVal = c.ToString();
        	
        	return retVal;
        }
        
        public String getMonth(ICell c)
        {
        	
        	String retVal = "";
        	if (c != null)
        	{
        		if (c.CellType == CellType.Numeric){
        			if (DateUtil.IsCellDateFormatted(c)){
        				DateTime dt = c.DateCellValue;
        				retVal = months[dt.Month-1];
        			}
        			else
        				retVal = c.NumericCellValue.ToString();
        		}
        		else if (c.CellType == CellType.String)
        			retVal = getCellValue(c);
        	}
        	
        	return retVal;
        }
        
        public String getCellDateValue(ICell c)
        {
        	String retVal = "";
        	if (c != null)
        	{
        		if (c.CellType == CellType.Numeric){
        			if (DateUtil.IsCellDateFormatted(c)){
        				DateTime dt = c.DateCellValue;
        				retVal = (dt.Day<10?"0":"") + dt.Day.ToString() + "-";
        				retVal += (dt.Month<10?"0":"") + dt.Month.ToString() + "-";
        				retVal += dt.Year;
        			}
        			else
        				retVal = c.NumericCellValue.ToString();
        		}
        		else if (c.CellType == CellType.String)
        			retVal = getCellValue(c);
        	}
        	
        	return retVal;
        }
	}
}