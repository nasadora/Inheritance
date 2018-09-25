using System;
using System.Data;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Drawing;



class ExcelUILanguageHelper : IDisposable
{
    private CultureInfo m_CurrentCulture;

    public ExcelUILanguageHelper()
    {
        // save current culture and set culture to en-US
        m_CurrentCulture = Thread.CurrentThread.CurrentCulture;
        Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
    }

    #region IDisposable Members

    public void Dispose()
    {
        // return to normal culture
        Thread.CurrentThread.CurrentCulture = m_CurrentCulture;
    }

    #endregion
}


class ExcelMyClass
{
    public virtual void SaveAs(
        [OptionalAttribute] Object Filename,
        [OptionalAttribute] Object FileFormat,
        [OptionalAttribute] Object Password,
        [OptionalAttribute] Object WriteResPassword,
        [OptionalAttribute] Object ReadOnlyRecommended,
        [OptionalAttribute] Object CreateBackup,
        [OptionalAttribute] Excel.XlSaveAsAccessMode AccessMode,
        [OptionalAttribute] Object ConflictResolution,
        [OptionalAttribute] Object AddToMru,
        [OptionalAttribute] Object TextCodepage,
        [OptionalAttribute] Object TextVisualLayout,
        [OptionalAttribute] Object Local
        )
    { }


    public virtual void Close(
            [OptionalAttribute] Object SaveChanges,
            [OptionalAttribute] Object Filename,
            [OptionalAttribute] Object RouteWorkbook)
    { }


    public void SaveFileAsXlsX(string sourceFile, string destinationFile)
     {
        using (new ExcelUILanguageHelper())
        {

            object missing = System.Reflection.Missing.Value;
            Excel.Application m_objExcel = new Excel.Application();
            Excel.Workbooks m_objBooks = (Excel.Workbooks)m_objExcel.Workbooks;

            try
            {

                m_objBooks.Open(sourceFile, 0, true, 5, null, null, true, Excel.XlPlatform.xlWindows, "\t", false,
                                false, 0, true, 10, 0);

                Excel.Workbook m_objBook = m_objExcel.ActiveWorkbook;
                //to add worksheets
                
                if (File.Exists(destinationFile)) File.Delete(destinationFile);

                m_objBook.SaveAs(destinationFile, //Excel.XlFileFormat.xlWorkbookNormal,
                    Excel.XlFileFormat.xlOpenXMLWorkbook,
                    missing, missing, false, false,
                Excel.XlSaveAsAccessMode.xlNoChange,
                Excel.XlSaveConflictResolution.xlLocalSessionChanges, missing, missing, missing, missing);
                //Microsoft.Office.Interop.
                m_objBook.Close(false, false, missing);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
                Console.WriteLine(ex.Source.ToString());
                Console.WriteLine(ex.StackTrace.ToString());

            }
            finally
            {
                m_objExcel.Quit();
            }
        }
    }

    public static void exportDsToExcelXML(DataSet dsSource, string fileName)
    {
        string header = "Test";
        string blank = "";
        //Excel.HeaderFooter = 

        System.IO.StreamWriter excelDoc;

        excelDoc = new System.IO.StreamWriter(fileName);
        const string startExcelXML = "<xml version>\r\n<Workbook " +
              "xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\r\n" +
              " xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n " +
              "xmlns:x=\"urn:schemas-    microsoft-com:office:" +
              "excel\"\r\n xmlns:ss=\"urn:schemas-microsoft-com:" +
              "office:spreadsheet\">\r\n <Styles>\r\n " +
              "<Style ss:ID=\"Default\" ss:Name=\"Normal\">\r\n " +
              "<Alignment ss:Vertical=\"Bottom\"/>\r\n <Borders/>" +
              "\r\n <Font/>\r\n <Interior/>\r\n <NumberFormat/>" +
              "\r\n <Protection/>\r\n </Style>\r\n " + 
              "<Style ss:ID=\"BoldHeader\">\r\n <Font " +
              "x:Family=\"Swiss\" ss:Bold=\"1\"/>\r\n </Style>\r\n" +             
              "<Style ss:ID=\"BoldColumn\">\r\n <Font " +
              "x:Family=\"Swiss\" ss:Bold=\"1\"/>\r\n </Style>\r\n " +
              "<Style     ss:ID=\"StringLiteral\">\r\n <NumberFormat" +
              " ss:Format=\"@\"/>\r\n </Style>\r\n <Style " +
              "ss:ID=\"Decimal\">\r\n <NumberFormat " +
              "ss:Format=\"0.0000\"/>\r\n </Style>\r\n " +
              "<Style ss:ID=\"Integer\">\r\n <NumberFormat " +
              "ss:Format=\"0\"/>\r\n </Style>\r\n <Style " +
              "ss:ID=\"DateLiteral\">\r\n <NumberFormat " +
              "ss:Format=\"mm/dd/yyyy;@\"/>\r\n </Style>\r\n " +
              "</Styles>\r\n ";
        const string endExcelXML = "</Workbook>";   


        /*
        <xml version>
        <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
        xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:x="urn:schemas-microsoft-com:office:excel"
        xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
        <Styles>
        <Style ss:ID="Default" ss:Name="Normal">
          <Alignment ss:Vertical="Bottom"/>
          <Borders/>
          <Font/>
          <Interior/>
          <NumberFormat/>
          <Protection/>
        </Style>
        <Style ss:ID="BoldColumn">
          <Font x:Family="Swiss" ss:Bold="1"/>
        </Style>
        <Style ss:ID="StringLiteral">
          <NumberFormat ss:Format="@"/>
        </Style>
        <Style ss:ID="Decimal">
          <NumberFormat ss:Format="0.0000"/>
        </Style>
        <Style ss:ID="Integer">
          <NumberFormat ss:Format="0"/>
        </Style>
        <Style ss:ID="DateLiteral">
          <NumberFormat ss:Format="mm/dd/yyyy;@"/>
        </Style>
        </Styles>
        <Worksheet ss:Name="Sheet1">
        </Worksheet>
        </Workbook>
        */

        excelDoc.Write(startExcelXML);
        
        
        foreach (DataTable dt in dsSource.Tables)
        {
            int rowCount = 0;
            int sheetCount = 2;           

            excelDoc.Write("<Worksheet ss:Name=\"" + dt.TableName.ToString() + "\">");            
            excelDoc.Write("<Table>");
            excelDoc.Write("<Row>");
           
            for (int x = 0; x < dt.Columns.Count; x++)
            {
                
                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">");
                excelDoc.Write(dt.Columns[x].ColumnName);
                excelDoc.Write("</Data></Cell>");
            }

            excelDoc.Write("</Row>");

            foreach (DataRow x in dt.Rows)
            {
                rowCount++;
                //if the number of rows is > 64000 create a new page to continue output

                if (rowCount == 64000)
                {
                    rowCount = 0;
                    sheetCount++;
                    excelDoc.Write("</Table>");
                    excelDoc.Write(" </Worksheet>");
                    //excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
                    excelDoc.Write("<Worksheet ss:Name=\"" + dt.TableName.ToString() + " part" + sheetCount + "\">");
                    excelDoc.Write("<Table>");
                }

                excelDoc.Write("<Row>"); //ID=" + rowCount + "
                //excelDoc.Write("nsdnaskds");
                //excelDoc.Write("");
                for (int y = 0; y < dt.Columns.Count; y++)
                {
                    System.Type rowType;
                    rowType = x[y].GetType();
                    switch (rowType.ToString())
                    {
                        case "System.String":
                            string XMLstring = x[y].ToString();
                            XMLstring = XMLstring.Trim();
                            XMLstring = XMLstring.Replace("&", "&amp;");
                            XMLstring = XMLstring.Replace(">", "&gt;");
                            XMLstring = XMLstring.Replace("<", "&lt;");

                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\">" +
                                           "<Data ss:Type=\"String\">");
                            excelDoc.Write(XMLstring);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.DateTime":
                            //Excel has a specific Date Format of YYYY-MM-DD followed by  

                            //the letter 'T' then hh:mm:sss.lll Example 2005-01-31T24:01:21.000

                            //The Following Code puts the date stored in XMLDate 

                            //to the format above

                            DateTime XMLDate = (DateTime)x[y];
                            string XMLDatetoString = ""; //Excel Converted Date

                            XMLDatetoString = XMLDate.Year.ToString() +
                                 "-" +
                                 (XMLDate.Month < 10 ? "0" +
                                 XMLDate.Month.ToString() : XMLDate.Month.ToString()) +
                                 "-" +
                                 (XMLDate.Day < 10 ? "0" +
                                 XMLDate.Day.ToString() : XMLDate.Day.ToString()) +
                                 "T" +
                                 (XMLDate.Hour < 10 ? "0" +
                                 XMLDate.Hour.ToString() : XMLDate.Hour.ToString()) +
                                 ":" +
                                 (XMLDate.Minute < 10 ? "0" +
                                 XMLDate.Minute.ToString() : XMLDate.Minute.ToString()) +
                                 ":" +
                                 (XMLDate.Second < 10 ? "0" +
                                 XMLDate.Second.ToString() : XMLDate.Second.ToString()) +
                                 ".000";
                            excelDoc.Write("<Cell ss:StyleID=\"DateLiteral\">" +
                                         "<Data ss:Type=\"DateTime\">");
                            excelDoc.Write(XMLDatetoString);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Boolean":
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\">" +
                                        "<Data ss:Type=\"String\">");
                            excelDoc.Write(x[y].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Int16":
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            excelDoc.Write("<Cell ss:StyleID=\"Integer\">" +
                                    "<Data ss:Type=\"Number\">");
                            excelDoc.Write(x[y].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Decimal":
                        case "System.Double":
                            excelDoc.Write("<Cell ss:StyleID=\"Decimal\">" +
                                  "<Data ss:Type=\"Number\">");
                            excelDoc.Write(x[y].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.DBNull":
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\">" +
                                  "<Data ss:Type=\"String\">");
                            excelDoc.Write("");
                            excelDoc.Write("</Data></Cell>");
                            break;
                        default:
                            throw (new Exception(rowType.ToString() + " not handled."));
                    }
                }
                excelDoc.Write("</Row>");
            }
            excelDoc.Write("</Table>");
            excelDoc.Write(" </Worksheet>");
        }
        excelDoc.Write(endExcelXML);
        excelDoc.Close();
    }
}
