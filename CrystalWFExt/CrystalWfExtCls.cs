using System;
using System.Configuration;
using System.Collections.Generic;
using System.Windows.Forms;

using LSSERVICEPROVIDERLib;
using System.Runtime.InteropServices;
using LSEXT;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

using System.Diagnostics;
using Oracle.ManagedDataAccess.Client;
using Patholab_Common;
using System.Reflection;

namespace CrystalWFExt
{
    [ComVisible(true)]
    [ProgId("CrystalWFExt.CrystalWfExtCls")]
    public class CrystalWfExtCls : IWorkflowExtension
    {

        #region private members

        #endregion

        public void Execute(ref LSExtensionParameters Parameters)
        {

            INautilusServiceProvider _sp;
            OracleConnection oraCon = null;
            OracleCommand cmd = null;
            ReportDocument cr = null;
            //Debugger.Launch();
            try
            {
                _sp = Parameters["SERVICE_PROVIDER"];
                string tableName = Parameters["TABLE_NAME"];

                int workstation = Parameters["WORKSTATION_ID"];
              //  Debugger.Launch();
                long wnid = Parameters["WORKFLOW_NODE_ID"];
                INautilusDBConnection _ntlsCon = null;
                if (_sp != null)
                {
                    _ntlsCon = _sp.QueryServiceProvider("DBConnection") as NautilusDBConnection;
                }
                else
                {
                    _ntlsCon = null;
                }
                if (_ntlsCon != null)
                {
                    // _username= dbConnection.GetUsername();
                    oraCon = GetConnection(_ntlsCon);
                    //set oracleCommand's connection
                    cmd = oraCon.CreateCommand();
                }

                var records = Parameters["RECORDS"];
                var recordId = records.Fields[tableName + "_ID"].Value;





                string sql =
                    string.Format(
                        "select parameter_2 from lims_sys.workflow_node where parent_id={0} and NAME='Comment' AND long_name='Path'",
                        wnid);
                string spath = GetPathFromConfig(workstation.ToString());
                if (string.IsNullOrEmpty(spath))
                {


                    cmd = new OracleCommand(sql, oraCon);
                    var path = cmd.ExecuteScalar();
                    spath = path.ToString();
                }
     


                sql =
              string.Format(
                  "select parameter_2 from lims_sys.workflow_node where parent_id={0} and NAME='Comment' AND long_name='Param'",
                  wnid);
                cmd.CommandText = sql;
                string param = (string)cmd.ExecuteScalar();



                string serverName;
                string nautilusUserName;
                string nautilusPassword;

                serverName = _ntlsCon.GetServerDetails();
                var IsProxy = _ntlsCon.GetServerIsProxy();
                if (IsProxy)
                {
                    nautilusUserName = "";
                    nautilusPassword = "";
                }
                else
                {
                    nautilusUserName = _ntlsCon.GetUsername();
                    nautilusPassword = _ntlsCon.GetPassword();
                }

                cr = new ReportDocument();
                TableLogOnInfo crTableLoginInfo;
                var crConnectionInfo = new ConnectionInfo();
                  
                cr.Load(spath);
                crConnectionInfo.ServerName = serverName;          
                if (IsProxy)
                {
                    crConnectionInfo.IntegratedSecurity = true;
                }
                else
                {
                    crConnectionInfo.UserID = nautilusUserName;
                    crConnectionInfo.Password = nautilusPassword;
                }

                cr.SetParameterValue(param, recordId.ToString());
                var CrTables = cr.Database.Tables;
                foreach (Table CrTable in CrTables)
                {
                    crTableLoginInfo = CrTable.LogOnInfo;
                    crTableLoginInfo.ConnectionInfo = crConnectionInfo;
                    CrTable.ApplyLogOnInfo(crTableLoginInfo);
                }

                cr.PrintToPrinter(1, true, 0, 0);




            }
            catch (Exception e)
            {
                MessageBox.Show("Error on CrystalWFExt.Execute: " + e.Message);
                Logger.WriteLogFile(e);
            }
            finally
            {
                if (cmd != null) cmd.Dispose();
                if (oraCon != null) oraCon.Close();


                if (cr != null)
                {
                    cr.Close();
                    cr.Dispose();
                }
            }


        }



        private string GetPathFromConfig(string workstation)
        {
            try
            {


                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                ExeConfigurationFileMap map = new ExeConfigurationFileMap();
                map.ExeConfigFilename = assemblyPath + ".config";
                Configuration cfg = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
                AppSettingsSection _appSettings = cfg.AppSettings;
                var key = _appSettings.Settings[workstation];
                if (key == null) return null;

                return key.Value;
            }
            catch (Exception ex)
            {

                return null;
            }

        }
        public OracleConnection GetConnection(INautilusDBConnection ntlsCon)
        {
            OracleConnection connection = null;
            if (ntlsCon != null)
            {
                //initialize variables
                string rolecommand;
                //try catch block
                try
                {

                    string connectionString;
                    string server = ntlsCon.GetServerDetails();
                    string user = ntlsCon.GetUsername();
                    string password = ntlsCon.GetPassword();

                    connectionString =
                        string.Format("Data Source={0};User ID={1};Password={2};", server, user, password);

                    var username = ntlsCon.GetUsername();
                    if (string.IsNullOrEmpty(username))
                    {
                        var serverDetails = ntlsCon.GetServerDetails();
                        connectionString = "User Id=/;Data Source=" + serverDetails + ";";
                    }

                    //create connection
                    connection = new OracleConnection(connectionString);

                    //open the connection
                    connection.Open();

                    //get lims user password
                    string limsUserPassword = ntlsCon.GetLimsUserPwd();

                    //set role lims user
                    if (limsUserPassword == "")
                    {
                        //lims_user is not password protected 
                        rolecommand = "set role lims_user";
                    }
                    else
                    {
                        //lims_user is password protected
                        rolecommand = "set role lims_user identified by " + limsUserPassword;
                    }

                    //set the oracle user for this connection
                    OracleCommand command = new OracleCommand(rolecommand, connection);

                    //try/catch block
                    try
                    {
                        //execute the command
                        command.ExecuteNonQuery();
                    }
                    catch (Exception f)
                    {
                        //throw the exeption
                        MessageBox.Show("Inconsistent role Security : " + f.Message);
                    }

                    //get session id
                    double sessionId = ntlsCon.GetSessionId();

                    //connect to the same session 
                    string sSql = string.Format("call lims.lims_env.connect_same_session({0})", sessionId);

                    //Build the command 
                    command = new OracleCommand(sSql, connection);

                    //execute the command
                    command.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    //throw the exeption
                    MessageBox.Show("Err At GetConnection: " + e.Message);
                }
            }
            return connection;
        }


    }
}
