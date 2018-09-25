using System.Data;
using System.Configuration;
using System;

class Reports_First : Reports
{
    public Reports_First(string connPrefix) : base(connPrefix)
    {
        ReportName = "X1 Report";
        ReportNameOne = "X2 Report";
        ReportNameTwo = "X3 Report";
        ReportGdiAd = "X4 Report";
        ReportNameSubject = "First Report";
        DestinationXLSfileName = "First_Report.xlsx";
        DestinationXMLfileName = "First_Report.xml";
    }

    public override DataSet X1_Report()
    {
        string commandText = "****** select statement";      

        DataSet ds = util.executeQuery(commandText, connectionPrefix);
        return ds;
    }

    public override DataSet X2_Report()
    {
        string commandText = "****** select statement";
        DataSet ds = util.executeQuery(commandText, connectionPrefix);
        return ds;
    }

    public override DataSet X3_Report()
    {
        string commandText = "****** select statement";

        DataSet ds = util.executeQuery(commandText, connectionPrefix);
        return ds;
    }

    public override DataSet X4_Report()
    {
        string commandText = "****** select statement";

        DataSet ds = util.executeQuery(commandText, connectionPrefix);
        return ds;
    }

    int CountGdiAdRecords()
    {
        string commandText = "****** select statement";

        return util.executeNonQuery(commandText, connectionPrefix);
    }

    public override DataSet readAdGdiFromTable()
    {
        mmandText = "****** select statement";
        DataSet ds = util.executeQueryWithData(commandText, connectionPrefix);
        return ds;
    }

    public override string GetBodyOfMail()
    {
        string DumpFileDate = "Data not available.";

        string commandText = "****** select statement";

        DataSet ds = util.executeQuery(commandText, connectionPrefix);
        if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
        {
            DumpFileDate = "Dumpfile Date: " + ds.Tables[0].Rows[0]["dumpfile_date"].ToString().Trim();
        }

        string body = "Test, \n\n";
        body += "Monthly Report for team feed. \n\n";
        body += "****************************************** \n";
        body += "* " + DumpFileDate + " * \n";
        body += "****************************************** \n\n\n";

        body += "If you have any queries, please let us know. \n\n\n";
        body += "Kind regards, \n";
        body += "Service Team \n";       
        body += AddDetailsToEmailBody();
        return body;
    }

    public override void DownloadGdiAdRecords()
    {
        if (ConfigurationManager.AppSettings["CountGdiAdRecords"].ToString().ToUpper().Trim() == "Y")
        CountGdiAdRecords();
       
    }
    
    public void RunReports()
    {
        Run();
    }
   
}
