using System;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml;

namespace HowToCreateExcelFileFromXml
{
  public partial class Form1 : Form
  {
    public Form1()
    {
      InitializeComponent();
    }

    private void button1_Click(object sender, EventArgs e)
    {
      Excel.Application xlApp;
      Excel.Workbook xlWorkBook;
      Excel.Worksheet xlWorkSheet;
      object misValue = System.Reflection.Missing.Value;

      DataSet ds = new DataSet();
      XmlReader xmlFile;
      //int i = 0;
      //int j = 0;

      xlApp = new Excel.ApplicationClass();
      xlWorkBook = xlApp.Workbooks.Add(misValue);
      xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

      // C:\Users\mvannortwick\Documents\Visual Studio 2008\Projects\howotcreateexcelfilefromxml\HowToCreateExcelFileFromXml\HowToCreateExcelFileFromXml\bin\Debug\product.xml
      // Support Document: "C:\Users\mvannortwick\Documents\Visual Studio 2008\Projects\howotcreateexcelfilefromxml\HowToCreateExcelFileFromXml\HowToCreateExcelFileFromXml\bin\Debug\product.xlsx"

      xmlFile = XmlReader.Create("Product.xml", new XmlReaderSettings());
      ds.ReadXml(xmlFile);
        


      string strcolLotID = ds.Tables[0].Columns[0].ColumnName;
      string strStartLot_Id = ds.Tables[0].Columns[1].ColumnName;
      string strTestRecipe = ds.Tables[0].Columns[2].ColumnName;
      string strTermination = ds.Tables[0].Columns[3].ColumnName;
      string strLotId = ds.Tables[0].Rows[0].ItemArray[0].ToString();
      string strTestRecipeName = ds.Tables[0].Rows[0].ItemArray[2].ToString();
      string strnumberofhours = ds.Tables[0].Rows[0].ItemArray[3].ToString();

      string strslots_Id = ds.Tables[1].Columns[0].ColumnName;
      string strcolStartLot_Id = ds.Tables[1].Columns[1].ColumnName;
      string strSlots = ds.Tables[1].Rows[0].ItemArray[0].ToString();
      string striii = ds.Tables[1].Rows[0].ItemArray[1].ToString();

      string strcolSlot_Id = ds.Tables[2].Columns[0].ColumnName;
      string strcolID = ds.Tables[2].Columns[1].ColumnName;
      string struu = ds.Tables[2].Columns[2].ColumnName;
      string strSlot1 = ds.Tables[2].Rows[0].ItemArray[1].ToString();

      string strx = ds.Tables[3].TableName;
      string strBoardd_Id = ds.Tables[3].Columns[0].ColumnName;
      string strID = ds.Tables[3].Columns[1].ColumnName;
      string strSlot_Id = ds.Tables[3].Columns[2].ColumnName;
      string strBoard1 = ds.Tables[3].Rows[0].ItemArray[1].ToString();

      string strPosition = ds.Tables[4].TableName;
      string strUnitID = ds.Tables[4].Columns["UnnnitXX1"].ColumnName;
      string strLocationID = ds.Tables[4].Columns[1].ColumnName;
      string stryy = ds.Tables[4].Columns[2].ColumnName;
      string str11 = ds.Tables[4].Columns[3].ColumnName;
      string str22 = ds.Tables[4].Columns[4].ColumnName;
      string str33 = ds.Tables[4].Columns[5].ColumnName;
      string str44 = ds.Tables[4].Columns[6].ColumnName;
      string str74 = ds.Tables[4].Columns[7].ColumnName;
      string str48 = ds.Tables[4].Columns[8].ColumnName;
      string str49 = ds.Tables[4].Columns[9].ColumnName;
      string strunit11 = ds.Tables[4].Rows[0].ItemArray[0].ToString();

      int nIndex = 0;
      foreach (DataRow dr in ds.Tables[4].Rows)
      {
          textBox1.AppendText(dr.ItemArray[nIndex].ToString() + "\n");
          textBox1.AppendText(dr.ItemArray[nIndex + 1].ToString() + "\n");
          nIndex += 2;
      }

      int nRow = 1;
      for (int nTableCounter = 0; nTableCounter < ds.Tables.Count; nTableCounter++)
      {
        for (int nRowCounter = 0; nRowCounter < ds.Tables[nTableCounter].Rows.Count; nRowCounter++)
        {
          for (int nItemArrayCounter = 0; nItemArrayCounter < ds.Tables[nTableCounter].Rows[nRowCounter].ItemArray.Length ; nItemArrayCounter++)
          {
            xlWorkSheet.Cells[nRowCounter + nRow, nItemArrayCounter + 1] = ds.Tables[nTableCounter].Rows[nRowCounter].ItemArray[nItemArrayCounter].ToString();
          }
        }
        nRow += ds.Tables[nTableCounter].Rows.Count;

        for (int nColumnCounter = 0; nColumnCounter < ds.Tables[nTableCounter].Columns.Count; nColumnCounter++)
        {
          xlWorkSheet.Cells[nRow, nColumnCounter + 1] = ds.Tables[nTableCounter].Columns[nColumnCounter].ColumnName;
        }

        nRow++;
      }

      // C:\Users\monte\Documents\xml2excel.xls
      xlWorkBook.SaveAs("xml2excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
      xlWorkBook.Close(true, misValue, misValue);
      xlApp.Quit();

      releaseObject(xlApp);
      releaseObject(xlWorkBook);
      releaseObject(xlWorkSheet);

      MessageBox.Show("Done .. ");
    }

    private void releaseObject(object obj)
    {
      try
      {
        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        obj = null;
      }
      catch (Exception ex)
      {
        obj = null;
      }
      finally
      {
        GC.Collect();
      }
    } 

  }
}
