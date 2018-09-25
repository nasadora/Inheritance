using System;
using System.Data;
using System.Data.OracleClient;
using System.Collections;
using ActiveDirectory;


public class OracleSource : IDisposable
{
    private OracleConnection _conn;
    private OracleTransaction _tr;
    private bool tr = false;

    public OracleSource()
    {
        _conn = new OracleConnection();
        _conn.ConnectionString = getConnectionString();
        _conn.Open();

    }

    public OracleSource(string connectionPrefix)
    {
        _conn = new OracleConnection();
        _conn.ConnectionString = getConnectionString(connectionPrefix);
        _conn.Open();
    }


    private OracleSource(string connectionPrefix, bool transactionEnabled)
    {
        _conn = new OracleConnection();
        _conn.ConnectionString = getConnectionString(connectionPrefix);
        _conn.Open();
        if (transactionEnabled)
        {
            BeginTransaction();
            tr = true;
        }
    }

    ~OracleSource()
    {
        EndTransaction();
        if (_conn.State.Equals(ConnectionState.Open)) _conn.Close();
        _conn.Dispose();
        OracleConnection.ClearPool(_conn);
        OracleConnection.ClearAllPools();

    }


    private void BeginTransaction()
    {
        _tr = _conn.BeginTransaction();
    }

    public void EndTransaction()
    {
        if (tr) _tr.Commit();
    }

    public void Dispose()
    {
        EndTransaction();
        if (_conn.State.Equals(ConnectionState.Open)) _conn.Close();
        _conn.Dispose();




        //OracleConnection.ClearPool(_conn);
        //OracleConnection.ClearAllPools();

        GC.SuppressFinalize(this);   //object

    }

    //uses configuration data to form a string that can be used to connect to a database
    private String getConnectionString()
    {
        return null;
    }

    //creates a condition wiht multiple statements connected by " OR "
    public string CreateCondition(string[] values, string columnName)
    {
        string condition = "";
        foreach (string currVal in values)
        {
            if (condition != "")
                condition += " OR ";
            condition += columnName + "=" + SecureOracleValue(currVal);
        }
        return condition;
    }

    private String getConnectionString(string connectionPrefix) //overloaded method
    {
        String encryptedPassword = System.Configuration.ConfigurationSettings.AppSettings[connectionPrefix + "EncryptedPassword"].ToString();
        String decryptedPassword = SecurePasswordHandling.Decrypt(encryptedPassword);
        String ConnectionString = "user id=" + System.Configuration.ConfigurationSettings.AppSettings[connectionPrefix + "User"] +
                                "; data source=" + System.Configuration.ConfigurationSettings.AppSettings[connectionPrefix + "DataSource"] +
                                "; password=" + decryptedPassword;
        return ConnectionString;
    }



    //returns value form a sequence
    public String getSequence(String sequenceName)
    {
        String sequenceValue = "";
        DataSet ds = executeQuery("SELECT " + sequenceName + ".nextval AS sq_value FROM DUAL");
        sequenceValue = (ds.Tables[0].Rows[0].ItemArray.GetValue(0)).ToString();

        return sequenceValue;
    }

    //creates an SQL insert statement into tableName using values from "data" hashtable
    public int CreateAndExecuteInsertQuery(string tableName, Hashtable data)
    {
        string columns = "", values = "";
        Hashtable queryParams = new Hashtable();
        //extract column names and values from data provided
        foreach (System.Collections.DictionaryEntry currentEntry in data)
        {
            string value = (string)(currentEntry.Value);
            if (columns != "")
            {
                columns += ", ";
                values += ", ";
            }
            columns += currentEntry.Key.ToString();
            //treat "sysdate" as function
            if (value == "sysdate" || value == "current_date")
            {
                values += value;
            }
            else
            {
                values += ":" + currentEntry.Key.ToString();
                queryParams[":" + currentEntry.Key.ToString()] = value;
            }
        }
        string query = "INSERT INTO " + tableName + "(" + columns + ") VALUES(" + values + ")";
        return executeNonQuery(query, queryParams);
    }

    //creates and executes an update on the tableName using data for column names and values
    public int CreateAndExecuteUpdateQuery(string tableName, Hashtable data, string condition)
    {
        string changes = "";
        Hashtable queryParams = new Hashtable();
        //extract column names and values from data provided
        foreach (System.Collections.DictionaryEntry currentEntry in data)
        {
            string value = (string)(currentEntry.Value);
            if (changes != "")
            {
                changes += ", ";
            }
            changes += currentEntry.Key.ToString() + "=";
            //treat "sysdate" as function
            if (value == "sysdate" || value == "current_date")
            {
                changes += value;
            }
            else
            {
                changes += ":" + currentEntry.Key.ToString();
                queryParams[":" + currentEntry.Key.ToString()] = value;
            }
        }
        string query = "UPDATE " + tableName + " SET " + changes + " WHERE " + condition;
        int result = executeNonQuery(query, queryParams);

        return result;
    }

    //an overload to executeQuery without parameters
    public DataSet executeQuery(String commandText)
    {
        return executeQuery(commandText, null);
    }
    public DataSet executeQueryWithData(String commandText)
    {
        return executeQueryWithData(commandText, null);
    }

    public bool executeQueryWithResult(String commandText)
    {
        return executeNonQueryWithResult(commandText, null);
    }

    public DataSet executeQueryWithData(String commandText, Hashtable param)
    {
        DataSet ds = new DataSet();
        OracleCommand oraCommand = new OracleCommand();
        OracleDataAdapter oda = new OracleDataAdapter();
        try
        {           
            oraCommand.Connection = _conn;
            oraCommand.CommandText = commandText;

            if (param != null)
            {
                foreach (DictionaryEntry e in param)
                {
                    OracleParameter p = new OracleParameter();
                    p.ParameterName = e.Key.ToString();
                    p.DbType = System.Data.DbType.AnsiString;
                    p.OracleType = OracleType.VarChar;

                    if (String.Equals(e.Value.ToString(), "null"))
                    {
                        p.Value = System.DBNull.Value;
                    }
                    else p.Value = e.Value.ToString();

                    oraCommand.Parameters.Add(p);
                    p = null;
                }
            }

            //if (_conn.State.Equals(ConnectionState.Closed))
            //{ _conn.Open(); };

            oraCommand.CommandType = CommandType.Text;
            //oraCommand.BindByName = true;
            oda.SelectCommand = oraCommand;
            oda.Fill(ds);           
            
            foreach (DataTable dt in ds.Tables)
            {
                foreach (DataRow x in dt.Rows)
                {
                    int rowCount = 0;
                    string mid = (x["XID"].ToString());
                    string admn = (x["XDOMAIN"].ToString());
                    string aact = (x["XACTIVE"].ToString());
                   
                    Active_Directory ad = new Active_Directory();
                    ad.getAdInfo(mid, admn, aact);
                    string a = ad.adAttribute;
                    if (a != aact)
                    {
                        string tableName = "ADMIN.TBL_XXXXXX";
                        CreateAndExecuteUpdateQueryWithData(tableName, a, mid);
                    }
                    rowCount++;
                }
                            
            }
        }

        catch (System.Data.OracleClient.OracleException ex)
        {
            Log logFile = new Log();
            logFile.LogMessageToFile("\n");
            logFile.LogMessageToFile("ERROR|" + ex.ToString() + "\n\n" + " The query was: " + oraCommand.CommandText + "\n\n\n");
            if (oraCommand.Parameters.Count > 0)
            {

                logFile.LogMessageToFile("Parameters number: " + oraCommand.Parameters.Count.ToString());
                for (int i = 0; i < oraCommand.Parameters.Count; i++)
                {
                    logFile.LogMessageToFile("Parameter's name: " + oraCommand.Parameters[i].ParameterName.ToString()
                        + "parameter's value: " + oraCommand.Parameters[i].Value.ToString());
                }
                logFile.LogMessageToFile("ERROR|" + "StackTrace: " + ex.StackTrace.ToString());
            }
            logFile.LogMessageToFile(ex.ToString());
            Console.WriteLine(ex.ToString());
            if (tr)
            {
                _tr.Rollback();
                tr = false;
            }


        }
        catch (Exception ex)
        {
            Log logFile = new Log();
            logFile.LogMessageToFile("\n");
            logFile.LogMessageToFile("ERROR|" + ex.ToString() + "\n\n" + " The query was: " + oraCommand.CommandText + "\n\n\n");
            if (oraCommand.Parameters.Count > 0)
            {

                logFile.LogMessageToFile("Parameters number: " + oraCommand.Parameters.Count.ToString());
                for (int i = 0; i < oraCommand.Parameters.Count; i++)
                {
                    logFile.LogMessageToFile("Parameter's name: " + oraCommand.Parameters[i].ParameterName.ToString()
                        + "parameter's value: " + oraCommand.Parameters[i].Value.ToString());
                }
            }
            logFile.LogMessageToFile("ERROR|" + "StackTrace: " + ex.StackTrace.ToString());
            logFile.LogMessageToFile(ex.ToString());

            Console.WriteLine(ex.ToString());
            if (tr)
            {
                _tr.Rollback();
                tr = false;
            }
        }

        return ds;
    }
    public int CreateAndExecuteUpdateQueryWithData(string tableName, string data, string condition)
    {        
        string query = "UPDATE " + tableName + " SET ADACTIVE = '" + data + "' WHERE KEYUID = '" + condition + "'";
        int result = executeNonQuery(query, null);     

        return result;
    }

    //executes a query on an Oracle database and returns result in a DataSet
    public DataSet executeQuery(String commandText, Hashtable param)
    {
        DataSet ds = new DataSet();
        OracleCommand oraCommand = new OracleCommand();
        OracleDataAdapter oda = new OracleDataAdapter();
        try
        {           
            oraCommand.Connection = _conn;
            oraCommand.CommandText = commandText;

            if (param != null)
            {
                foreach (DictionaryEntry e in param)
                {
                    OracleParameter p = new OracleParameter();
                    p.ParameterName = e.Key.ToString();
                    p.DbType = System.Data.DbType.AnsiString;
                    p.OracleType = OracleType.VarChar;

                    if (String.Equals(e.Value.ToString(), "null"))
                    {
                        p.Value = System.DBNull.Value;
                    }
                    else p.Value = e.Value.ToString();

                    oraCommand.Parameters.Add(p);
                    p = null;
                }
            }

            //if (_conn.State.Equals(ConnectionState.Closed))
            //{ _conn.Open(); };

            oraCommand.CommandType = CommandType.Text;
            //oraCommand.BindByName = true;
            oda.SelectCommand = oraCommand;
            oda.Fill(ds);
          
        }

        catch (System.Data.OracleClient.OracleException ex)
        {
            Log logFile = new Log();
            logFile.LogMessageToFile("\n");
            logFile.LogMessageToFile("ERROR|" + ex.ToString() + "\n\n" + " The query was: " + oraCommand.CommandText + "\n\n\n");
            if (oraCommand.Parameters.Count > 0)
            {

                logFile.LogMessageToFile("Parameters number: " + oraCommand.Parameters.Count.ToString());
                for (int i = 0; i < oraCommand.Parameters.Count; i++)
                {
                    logFile.LogMessageToFile("Parameter's name: " + oraCommand.Parameters[i].ParameterName.ToString()
                        + "parameter's value: " + oraCommand.Parameters[i].Value.ToString());
                }
                logFile.LogMessageToFile("ERROR|" + "StackTrace: " + ex.StackTrace.ToString());
            }
            logFile.LogMessageToFile(ex.ToString());
            Console.WriteLine(ex.ToString());
            if (tr)
            {
                _tr.Rollback();
                tr = false;
            }


        }
        catch (Exception ex)
        {
            Log logFile = new Log();
            logFile.LogMessageToFile("\n");
            logFile.LogMessageToFile("ERROR|" + ex.ToString() + "\n\n" + " The query was: " + oraCommand.CommandText + "\n\n\n");
            if (oraCommand.Parameters.Count > 0)
            {

                logFile.LogMessageToFile("Parameters number: " + oraCommand.Parameters.Count.ToString());
                for (int i = 0; i < oraCommand.Parameters.Count; i++)
                {
                    logFile.LogMessageToFile("Parameter's name: " + oraCommand.Parameters[i].ParameterName.ToString()
                        + "parameter's value: " + oraCommand.Parameters[i].Value.ToString());
                }
            }
            logFile.LogMessageToFile("ERROR|" + "StackTrace: " + ex.StackTrace.ToString());
            logFile.LogMessageToFile(ex.ToString());

            Console.WriteLine(ex.ToString());
            if (tr)
            {
                _tr.Rollback();
                tr = false;
            }
        }

        return ds;
    }


    //encloses values in apostrophes
    //escapes existing apostrophes
    public static string SecureOracleValue(string value)
    {
        if (value == null)
            return "''";
        //do not enclose oracle functions in apostrohes
        if (value == "sysdate" || value == "current_date")
            return value;
        value = value.Replace("'", "''");
        return "'" + value + "'";
    }


    //an overload to executeNonQuery without parameters
    public int executeNonQuery(String commandText)
    {
        return executeNonQuery(commandText, null);
    }

    public string executeString()
    {
        return executeString();
    }
    //executes a command that modifies Oracle data (as opposed to just selecting data)
    //returns number of rows affected by the nonQuery
    public int executeNonQuery(String commandText, Hashtable param)
    {

        OracleCommand oraCommand = new OracleCommand();
        OracleDataAdapter oda = new OracleDataAdapter();
        int rowsAffected = 0; //number to check if command was executed

        try
        {
            /*  if (tr)
              {
                   oraCommand.Transaction = _tr;
              }*/

            _tr = oraCommand.Transaction;
            oraCommand.Connection = _conn;
            oraCommand.CommandText = commandText;

            if (param != null)
            {
                string value = "";
                foreach (DictionaryEntry e in param)
                {
                    OracleParameter p = new OracleParameter();
                    p.OracleType = OracleType.VarChar;

                    p.ParameterName = e.Key.ToString();
                    value = (e.Value.ToString()).Trim();
                    p.Value = value;
                    if (value.Length >= 4000)
                    {
                        // p.OracleDbType = OracleDbType.NClob;
                        p.OracleType = OracleType.NClob;
                    }
                    oraCommand.Parameters.Add(p);
                    p = null;
                }
            }


            oraCommand.CommandType = CommandType.Text;
            oda.SelectCommand = oraCommand;
            //oraCommand.BindByName = true;
            rowsAffected = oraCommand.ExecuteNonQuery();
            Console.WriteLine(rowsAffected + " has been dowloaded to table");
        }

        catch (Exception ex)
        {
            Log logFile = new Log();
            logFile.LogMessageToFile("\n");
            logFile.LogMessageToFile("ERROR|" + ex.Message + " The non-query was: " + oraCommand.CommandText);
            if (oraCommand.Parameters.Count > 0)
            {

                logFile.LogMessageToFile("Parameters number: " + oraCommand.Parameters.Count.ToString());
                for (int i = 0; i < oraCommand.Parameters.Count; i++)
                {
                    logFile.LogMessageToFile("Parameter's name: " + oraCommand.Parameters[i].ParameterName.ToString()
                        + "parameter's value: " + oraCommand.Parameters[i].Value.ToString());
                }
            }
            logFile.LogMessageToFile("ERROR|" + "StackTrace: " + ex.StackTrace.ToString());
            logFile.LogMessageToFile(ex.ToString());
            Console.WriteLine(ex.ToString());
            if (tr)
            {
                _tr.Rollback();
                tr = false;
            }
        }
        return rowsAffected;
    }

    public string executeString(string connectionPrefix)
    {
        string returnValue = string.Empty;
        try
        {
            using (OracleConnection chameleonConn = new OracleConnection())
            {

                chameleonConn.Open();
                using (OracleCommand cmd = chameleonConn.CreateCommand())  // updated code from Priv team 
                {
                    cmd.CommandText = "SELECT keyuid from TBL_UNDELETED_AD_ACCT_STATUS";
                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            returnValue = reader.GetString(0);   // if the group is listed the hastrows will be true, else its false. 
                        }
                    }
                }

            }

        }
        catch
        {
            return "error"; // restriction failed or could not find the group so retrun true
        }

        return returnValue;
    }

    public bool executeNonQueryWithResult(String commandText, Hashtable param)
    {

        OracleCommand oraCommand = new OracleCommand();
        OracleDataAdapter oda = new OracleDataAdapter();
        int rowsAffected = 0;

        bool result = false;

        try
        {
            /*  if (tr)
              {
                   oraCommand.Transaction = _tr;
              }*/

            _tr = oraCommand.Transaction;
            oraCommand.Connection = _conn;
            oraCommand.CommandText = commandText;

            if (param != null)
            {
                string value = "";
                foreach (DictionaryEntry e in param)
                {
                    OracleParameter p = new OracleParameter();
                    p.OracleType = OracleType.VarChar;

                    p.ParameterName = e.Key.ToString();
                    value = (e.Value.ToString()).Trim();
                    p.Value = value;
                    if (value.Length >= 4000)
                    {
                        // p.OracleDbType = OracleDbType.NClob;
                        p.OracleType = OracleType.NClob;
                    }
                    oraCommand.Parameters.Add(p);
                    p = null;
                }
            }


            //if (_conn.State.Equals(ConnectionState.Closed))
            //{ _conn.Open(); }


            oraCommand.CommandType = CommandType.Text;
            oda.SelectCommand = oraCommand;
            //oraCommand.BindByName = true;
            rowsAffected = oraCommand.ExecuteNonQuery();
            result = true;
        }

        catch (Exception ex)
        {
            Log logFile = new Log();
            logFile.LogMessageToFile("\n");
            logFile.LogMessageToFile("ERROR|" + ex.Message + " The non-query was: " + oraCommand.CommandText);
            if (oraCommand.Parameters.Count > 0)
            {

                logFile.LogMessageToFile("Parameters number: " + oraCommand.Parameters.Count.ToString());
                for (int i = 0; i < oraCommand.Parameters.Count; i++)
                {
                    logFile.LogMessageToFile("Parameter's name: " + oraCommand.Parameters[i].ParameterName.ToString()
                        + "parameter's value: " + oraCommand.Parameters[i].Value.ToString());
                }
            }
            logFile.LogMessageToFile("ERROR|" + "StackTrace: " + ex.StackTrace.ToString());
            logFile.LogMessageToFile(ex.ToString());
            Console.WriteLine(ex.ToString());
            if (tr)
            {
                _tr.Rollback();
                tr = false;
            }
        }
        return result;
    }

    public string executeTruncSP(string tblName)
    {
        OracleCommand oraCommand = new OracleCommand();
        try
        {
            //OracleConnection trCon = new OracleConnection(_conn);
            oraCommand.Connection = _conn;
            //trCon.Open();
            //_conn.Open();
            OracleCommand trComm = new OracleCommand("USERADMIN.TRUNCATEUNDELETEDRPT", _conn);

            trComm.CommandType = CommandType.StoredProcedure;
            OracleParameter prm1 = new OracleParameter("tblname", OracleType.VarChar); //OracleDbType.Varchar2);
            prm1.Direction = ParameterDirection.Input;
            prm1.Value = tblName;
            trComm.Parameters.Add(prm1);

            OracleParameter prm2 = new OracleParameter("success", OracleType.Number); //OracleDbType.Varchar2);
            prm2.Direction = ParameterDirection.Output;
            trComm.Parameters.Add(prm2);
            trComm.ExecuteNonQuery();
            //trCon.Close();
            //_conn.Close();
            return "table has been truncated";
        }
        catch (Exception e)
        {
            return e.ToString();
        }

    }

    public DataTable readOracleToDatabtable(DataTable gdiAdTable, string commandTxt)
    {
        // put in gdiAdTable to DataTable
        int tblRow = 0;

        OracleCommand oraCommand = new OracleCommand();        
        oraCommand.CommandText = commandTxt;

                try
                {
                    using (OracleDataReader reader = oraCommand.ExecuteReader())
                    {

                        // Check for null value from the database
                        while (reader.Read())
                        {

                            {
                                gdiAdTable.Rows.Add("");

                                try
                                {

                                    gdiAdTable.Rows[tblRow]["GDI_KEYUID"] = reader.GetOracleString(0);
                                    gdiAdTable.Rows[tblRow]["GDI_STATUS"] = reader.GetOracleString(1);
                                    gdiAdTable.Rows[tblRow]["AD_DOMAIN"] = reader.GetOracleString(2);
                                    gdiAdTable.Rows[tblRow]["AD_ACCTACTIVE"] = reader.GetDateTime(3);
                                    gdiAdTable.Rows[tblRow]["AD_LAST_LOGIN"] = reader.GetDateTime(4);
                                    gdiAdTable.Rows[tblRow]["AD_ACCT_EXPIRY"] = reader.GetDateTime(5);
                                    gdiAdTable.Rows[tblRow]["AD_DUMPFILE_DATE"] = reader.GetDateTime(6);
                                   
                                    tblRow++;

                                }
                                catch (Exception e)
                                {
                                    // do nothing 

                                    DateTime datetime = DateTime.Now;
                                    string MDYformat = "dd/MMM/yyyy hh:mm:ss";
                                    string justDateN = datetime.ToString(MDYformat);
                                    
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    DateTime datetime = DateTime.Now;
                    string MDYformat = "dd/MMM/yyyy hh:mm:ss";
                    string justDateN = datetime.ToString(MDYformat);
                   
                }          
        
        return gdiAdTable;
    }

}
