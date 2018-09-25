using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Security.Principal;
using System.Configuration;
using ActiveDirectory;

public abstract class Reports
{
    protected AppConfig app = new AppConfig();
    protected string connectionPrefix;
    protected Utilities util = new Utilities();    

    public Reports(string connPrefix)
    {
        connectionPrefix = connPrefix;
    }

    private string _reportGdiAd;
    private string _reportName;
    private string _reportNameOne;
    private string _reportNameTwo;
    private string _reportNameSubject;

    public string ReportNameSubject
    {   get { return _reportNameSubject; }
        set { _reportNameSubject = value; }
    }
    public string ReportGdiAd
    {
        get { return _reportGdiAd; }
        set { _reportGdiAd = value; }
    }
    public string ReportName
    {
        get { return _reportName; }
        set { _reportName = value; }
    }
    public string ReportNameOne
    {
        get { return _reportNameOne; }
        set { _reportNameOne = value; }
    }
    public string ReportNameTwo
    {
        get { return _reportNameTwo; }
        set { _reportNameTwo = value; }
    }

    private string _destinationXMLfile;
    public string DestinationXMLfileName
    {
        get { return _destinationXMLfile; }
        set { _destinationXMLfile = value; }
    }

    private string _destinationXLSfile;
    public string DestinationXLSfileName
    {
        get { return _destinationXLSfile; }
        set { _destinationXLSfile = value; }
    }

    public abstract string GetBodyOfMail();
    public abstract DataSet X1_Report();
    public abstract DataSet X2_Report();  
    public abstract DataSet X3_Report();
    public abstract DataSet X4_Report();
    public abstract void DownloadGdiAdRecords();
    public abstract DataSet readAdGdiFromTable();    
  
    public void ExportDataToExcelXMLFile(string destinationFile)
    {

        DataSet dsReport = X1_Report();
        DataSet dsReportOne = X2_Report();
        DataSet dsReportTwo = X3_Report();
        DataSet dsReportThree = X4_Report();

        if (dsReport.Tables.Count == 0) dsReport.Tables.Add();
        if (dsReportOne.Tables.Count == 0) dsReportOne.Tables.Add();
        if (dsReportTwo.Tables.Count == 0) dsReportTwo.Tables.Add();
        if (dsReportThree.Tables.Count == 0) dsReportThree.Tables.Add();

        dsReport.Tables[0].TableName = ReportName;
        dsReportOne.Tables[0].TableName = ReportNameOne;
        dsReportTwo.Tables[0].TableName = ReportNameTwo;
        dsReportThree.Tables[0].TableName = ReportGdiAd;

        DataSet ds = new DataSet();
        ds.Tables.Add(dsReport.Tables[0].Copy());
        ds.Tables.Add(dsReportOne.Tables[0].Copy());
        ds.Tables.Add(dsReportTwo.Tables[0].Copy());
        ds.Tables.Add(dsReportThree.Tables[0].Copy());

        ExcelMyClass.exportDsToExcelXML(ds, destinationFile);
    }    

    public void Run()
    {
        Console.WriteLine(System.String.Format("{0:G}", System.DateTime.Now) + "  *** START *** " + ReportNameSubject);
        Console.WriteLine();
        
        Console.WriteLine(System.String.Format("{0:G}", System.DateTime.Now) + "  1: Truncate TBL_XXXXXX ");
        util.executeTruncSP("TBL_XXXXXX", connectionPrefix);
      
        Console.WriteLine(System.String.Format("{0:G}", System.DateTime.Now) + "  2: Download GDI & AD Records");
        DownloadGdiAdRecords();
     
        Console.WriteLine(System.String.Format("{0:G}", System.DateTime.Now) + "  3: Read table TBL_XXXXXX");
        readAdGdiFromTable();  

        Console.WriteLine(System.String.Format("{0:G}", System.DateTime.Now) + "  4: Exporting Report to XML file");
        string destinationXMLfile = AppConfig.GetXMLFilePath() + _destinationXMLfile;
        ExportDataToExcelXMLFile(destinationXMLfile);

        Console.WriteLine(System.String.Format("{0:G}", System.DateTime.Now) + "  5: Converting XML file to XLSX");
        string sourceXMLFile = destinationXMLfile;
        string destinationXLSFile = AppConfig.GetXLSFilePath() + Utilities.AddSysdateToFileName(_destinationXLSfile);
        ExcelMyClass excel = new ExcelMyClass();

        excel.SaveFileAsXlsX(sourceXMLFile, destinationXLSFile);

        Console.WriteLine(System.String.Format("{0:G}", System.DateTime.Now) + "  6: Sending report via email to " + app.GetRecipients());

        Mail mail = new Mail();
        string subject = "SERVICE_REPORT: " + _reportNameSubject;

        string body = GetBodyOfMail();
        mail.SendMailWithAttachment(subject, body, destinationXLSFile);

        Console.WriteLine(System.String.Format("{0:G}", System.DateTime.Now) + "  7: Copying file(s) to " + AppConfig.GetReportFolderPath());

        string directoryPath = AppConfig.GetReportFolderPath();
        if (!Directory.Exists(directoryPath)) Directory.CreateDirectory(directoryPath);

        if (File.Exists(destinationXLSFile))
        {
            string destFile = directoryPath + Utilities.AddSysdateToFileName(_destinationXLSfile);
            try
            {
                if (File.Exists(destFile)) File.Delete(destFile);
                if (File.Exists(destinationXLSFile)) File.Move(destinationXLSFile, destFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.Source);
                Console.WriteLine(ex.StackTrace);
            }
        }

        if (AppConfig.GetDeleteXMLFile().ToUpper().Trim() == "Y")
        {
            try
            {
                if (File.Exists(sourceXMLFile)) File.Delete(sourceXMLFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.Source);
                Console.WriteLine(ex.StackTrace);
            }
        }
        else
        {
            if (File.Exists(sourceXMLFile))
            {
                string descFile = directoryPath + Utilities.AddSysdateToFileName(_destinationXMLfile);
                try
                {
                    if (File.Exists(descFile)) File.Delete(descFile);
                    File.Move(sourceXMLFile, descFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Console.WriteLine(ex.Source);
                    Console.WriteLine(ex.StackTrace);
                }
            }
        }
        Console.WriteLine(System.String.Format("{0:G}", System.DateTime.Now) + "  *** END ***");
    }    

    public string AddDetailsToEmailBody()
    {
        string body = "";
        body += "############################################################\n";
        body += "# This email was generated automatically. \n";
        body += "# Sent at: " + System.String.Format("{0:G}", System.DateTime.Now) + " \n";
        body += "# Sent by account: " + WindowsIdentity.GetCurrent().Name.ToString() + " \n";
        body += "# Sent from machine: " + Environment.MachineName + " \n";
        body += "############################################################ \n";
        return body;

    } 


}



