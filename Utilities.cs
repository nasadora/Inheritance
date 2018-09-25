using System;
using System.Data;


public class Utilities
{

    public DataSet executeQuery(string commandText, string connectionPrefix)
    {
        DataSet ds = new DataSet();
        using (OracleSource oracleSource = new OracleSource(connectionPrefix))
        {
            ds = oracleSource.executeQuery(commandText);
        }
        return ds;
    }

    public DataSet executeQueryWithData(string commandText, string connectionPrefix)
    {
        DataSet ds = new DataSet();
        using (OracleSource oracleSource = new OracleSource(connectionPrefix))
        {
            ds = oracleSource.executeQueryWithData(commandText);
        }
        return ds;
    }

    public int executeNonQuery(string commandText, string connectionPrefix)
    {
        int affectedRow = 0;
        using (OracleSource oracleSource = new OracleSource(connectionPrefix))
        {
            affectedRow = oracleSource.executeNonQuery(commandText);
        }
        return affectedRow;
    }

    public string executeString(string connectionPrefix)
    {
        string message = string.Empty;
        using (OracleSource oracleSource = new OracleSource(connectionPrefix))
        {
            message = oracleSource.executeString(connectionPrefix);
        }
        return message;
    }

    public string executeTruncSP(string tblName, string connectionPrefix)
    {
        string message = string.Empty;
        using (OracleSource oracleSource = new OracleSource(connectionPrefix))
        {
            message = oracleSource.executeTruncSP(tblName);
        }
        return message;
    }

    public static string AddSysdateToFileName(string fName)
    {
        string[] fileTable = fName.Split('.');
        string sYear = DateTime.Now.Year.ToString();
        string sMonth = String.Format("{0:MM}", DateTime.Now);
        string sDay = String.Format("{0:dd}", DateTime.Now);
        string date = sYear + sMonth + sDay;

        string fileName = fileTable[0] + "_" + date + "." + fileTable[1];
        return fileName;
    }

}
