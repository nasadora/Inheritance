
using System.Configuration;
using System.IO;

public class AppConfig
{
    public static string GetConstant(string constantName)
    {
        return System.Configuration.ConfigurationManager.AppSettings[constantName];
    }

    //returns constant split into an array
    public static string[] GetConstantArray(string constantName)
    {
        char separator = (GetConstant("ConstantFieldSeparator"))[0];
        string values = GetConstant(constantName);
        return values.Split(separator);
    }

    public string GetConnectionPrefix()
    {
        return ConfigurationManager.AppSettings["ConnectionPrefix"].ToString();
    }

    public string GetSendEmailsWithErrorValue()
    {
        return ConfigurationManager.AppSettings["SendEmailsWithError"].ToString();
    }

    public static string GetConsoleOutputValue()
    {
        return ConfigurationManager.AppSettings["ConsoleOutput"].ToString();
    }

    public string GetSendRemedyTicketsValue()
    {
        return ConfigurationManager.AppSettings["SendRemedyTickets"].ToString();
    }

    public string GetCheckDataInSharedArea()
    {
        return ConfigurationManager.AppSettings["CheckDataInSharedArea"].ToString();
    }

    public string GetDropOldDataFromSeUnixcreated() { return ConfigurationManager.AppSettings["DropOldDataFromSeUnixcreated"].ToString(); }

    public string GetCheckSoxServerList() { return ConfigurationManager.AppSettings["CheckServerList"].ToString(); }

    public string GetSharedAreaName() { return ConfigurationManager.AppSettings["SharedName"].ToString(); }

    public string GetSharedAreaUserName() { return ConfigurationManager.AppSettings["SharedAreaUserName"].ToString(); }

    public string GetSharedAreaPassword() { return SecurePasswordHandling.Decrypt(ConfigurationManager.AppSettings["SharedAreaPassword"].ToString()); }

    public string GetDayOfTheWeek() { return ConfigurationManager.AppSettings["DayOfTheWeek"].ToString(); }

    public string GetFileNameFormatDate() { return ConfigurationManager.AppSettings["FileNameFormatDate"].ToString(); }

    public string GetFileName() { return ConfigurationManager.AppSettings["FileName"].ToString(); }

    public string GetRecipients() { return ConfigurationManager.AppSettings["EmailRecipients"].ToString(); }

    public string GetCC() { return ConfigurationManager.AppSettings["EmailCC"].ToString(); }

    public string GetSmtpHost() { return ConfigurationManager.AppSettings["SmtpHost"].ToString(); }

    public string GetMailboxAdrress() { return ConfigurationManager.AppSettings["AccessComplianceMailboxAddress"].ToString(); }

    public string GetDisplaySender() { return ConfigurationManager.AppSettings["DisplayNameOfMailSender"].ToString(); }

    public string GetCodeRunsOnServer() { return ConfigurationManager.AppSettings["CodeRunsOnServer"].ToString(); }

    public string GetDateFormat() { return ConfigurationManager.AppSettings["DateFormat"].ToString(); }
    public string GetDateFormatNET() { return ConfigurationManager.AppSettings["DateFormat.NET"].ToString(); }

    //remedy
    public static string GetRemedyAssignedToGroup() { return ConfigurationManager.AppSettings["AssignedToGroup"].ToString(); }
    public string GetRemedyType() { return ConfigurationManager.AppSettings["Type"].ToString(); }
    public string GetRemedyCategory() { return ConfigurationManager.AppSettings["Category"].ToString(); }
    public string GetRemedySubCode() { return ConfigurationManager.AppSettings["SubCode"].ToString(); }
    public string GetRemedyRequesterLoginName() { return ConfigurationManager.AppSettings["RequesterLoginName"].ToString(); }
    public string GetRemedyWebServiceToken() { return ConfigurationManager.AppSettings["WebServiceToken"].ToString(); }
    public string GetRemedySubmitedBy() { return ConfigurationManager.AppSettings["SubmitedBy"].ToString(); }

    //Log
    public static string GetLogPath() { return ConfigurationManager.AppSettings["LogPath"].ToString(); }
    public static string GetLogFileName() { return ConfigurationManager.AppSettings["LogFileName"].ToString(); }


    //ExcelReports
    public static string GetXMLFilePath()
    {
        string path = ConfigurationManager.AppSettings["XMLFilePath"].ToString();
        if (path == ".") path = Directory.GetCurrentDirectory();
        if (!path.EndsWith("\\")) path += "\\";
        return path;
    }

    public static string GetXLSFilePath()
    {
        string path = ConfigurationManager.AppSettings["XLSFilePath"].ToString();
        if (path == ".") path = Directory.GetCurrentDirectory();
        if (!path.EndsWith("\\")) path += "\\";
        return path;
    }


    public static string GetReportFolderPath()
    {
        string path = ConfigurationManager.AppSettings["ReportFolderPath"].ToString();
        if (path == ".") path = Directory.GetCurrentDirectory();
        if (!path.EndsWith("\\")) path += "\\";
        return path;
    }

    public static string GetDeleteXMLFile() { return ConfigurationManager.AppSettings["deleteXMLFile"].ToString(); }

    public static string GetDestinationXMLfileName() { return ConfigurationManager.AppSettings["destinationXMLfileName"].ToString(); }

    public static string GetDestinationXLSfileName() { return ConfigurationManager.AppSettings["destinationXLSfileName"].ToString(); }

    public static string GetExcelReportMailSubject() { return ConfigurationManager.AppSettings["excelReportMailSubject"].ToString(); }

    public static string GetExcelReportMailBody() { return ConfigurationManager.AppSettings["excelReportMailBody"].ToString(); }


}

