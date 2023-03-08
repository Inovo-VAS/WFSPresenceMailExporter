using System;
using System.Collections.Generic;
using System.Configuration;
using System.Web;

namespace PresenceMailArchiver
{
    public class IISHandlerArchiver : IHttpHandler
    {
        private static string mailarchive = ConfigurationManager.ConnectionStrings["mailarchive"].ConnectionString;

        private static System.Data.Common.DbConnection sqlcn = dbconnection();
        private static System.Data.Common.DbConnection dbconnection()
        {
            try
            {
                var cndb = Microsoft.Data.SqlClient.SqlClientFactory.Instance.CreateConnection();
                cndb.ConnectionString = mailarchive;
                cndb.Open();
                LogMessage("info:", "dbconnection()", "db-connected");
                return (System.Data.Common.DbConnection)cndb;
            }
            catch (Exception ex)
            {
                LogMessage("err:", "dbconnection()", ex.Message);
            }
            return null;
        }
        /// <summary>
        /// You will need to configure this handler in the Web.config file of your 
        /// web and register it with IIS before being able to use it. For more information
        /// see the following link: https://go.microsoft.com/?linkid=8101007
        /// </summary>
        #region IHttpHandler Members

        public bool IsReusable
        {
            // Return false in case your Managed Handler cannot be reused for another request.
            // Usually this would be false in case you have some state information preserved per request.
            get { return true; }
        }

        public void ProcessRequest(HttpContext context)
        {
            var req = context.Request;
            var resp = context.Response;
            Dictionary<String, string> archiveRequestSettings = new Dictionary<string, string>();

            Dictionary<String, String> parameters = new Dictionary<string, string>();

            foreach (var item in req.QueryString.AllKeys)
            {
                parameters.Add(item.ToUpper(), req.QueryString[item]);

            }

            string archssettings = "";
            archiveRequestSettings["AGENTID"] = parameters.ContainsKey("AGENTID") ? parameters["AGENTID"] : ""; ;
            archiveRequestSettings["METAFIELD1NAME"] = "CUSTOMERID";
            archiveRequestSettings["METAFIELD1VALUE"] = parameters.ContainsKey("CUSTOMERID") ? parameters["CUSTOMERID"] : "0000000000000";
            archiveRequestSettings["INBOUNDMAILID"] = parameters.ContainsKey("INBOUNDMAILID") ? parameters["INBOUNDMAILID"] : "";

            foreach(var item in archiveRequestSettings)
            {
                archssettings += item.Key + ":" + item.Value + ",";

            }
            LogMessage("info", "[archiveRequestSettings] " + archssettings.Substring(0, archssettings.Length - 1));
            resp.ContentType = "text/plain";
            try
            {
                String validationError = "";
                if (validationError.Equals(""))
                {
                    validationError = this.requestMailArchive(archiveRequestSettings, new String[] { "AGENTID", "INBOUNDMAILID", "METAFIELD1VALUE" }, "PRSC_[AGENTID]_[METAFIELD1VALUE]_[MAILID]", "METAFIELD1".Split(",".ToCharArray()), true);
                }
                if (validationError.Equals(""))
                {
                    resp.Write("SUCCESSFUL");
                }
                else
                    resp.Write("UNSUCCESSFUL:" + validationError.Replace("\n", " "));
            }
            catch (Exception e)
            {
                resp.Write("ERROR:" + e.Message.Replace("\n", " "));

            }
        }

        #endregion

        public String requestMailArchive(Dictionary<string, string> archiveRequestSettings, String[] archiveMailIDLayout, String archiveFileMaskNameLayout, String[] requiredMetaFields, bool insertIntoDbDirect)
        {
            String validationError = "";

            if (archiveRequestSettings.ContainsKey("AGENTID"))
            {
                if (archiveRequestSettings["AGENTID"].Equals(""))
                {
                    validationError = "AGENT ID NOT PROVIDED,";
                }
            }
            else
            {
                validationError = "AGENT ID NOT PROVIDED,";
            }
            if (archiveRequestSettings.ContainsKey("INBOUNDMAILID"))
            {
                if (archiveRequestSettings["INBOUNDMAILID"].Equals(""))
                {
                    validationError = validationError + "INBOUNDMAILID ID NOT PROVIDED,";
                }
            }
            else
            {
                validationError = validationError + "INBOUNDMAILID ID NOT PROVIDED,";
            }

            foreach (String metaFieldRef in requiredMetaFields)
            {
                if (archiveRequestSettings.ContainsKey(metaFieldRef + "NAME"))
                {
                    if (archiveRequestSettings.ContainsKey(metaFieldRef + "VALUE"))
                    {
                        if (archiveRequestSettings[metaFieldRef + "VALUE"].Equals(""))
                        {
                            validationError = validationError + archiveRequestSettings[metaFieldRef+"NAME"] + " NOT PROVIDED,";
                        }
                    }
                    else
                    {
                        validationError = validationError + archiveRequestSettings[metaFieldRef+"NAME"] + " NOT PROVIDED,";
                    }
                }
            }
            if (validationError.Equals(""))
            {
                if ((sqlcn = sqlcn == null ? dbconnection() : sqlcn) != null)
                {
                    var sqlcmd = sqlcn.CreateCommand();
                    try
                    {
                        sqlcmd.CommandText = "INSERT INTO PTOOLS.MAILEXPORTSTAGING (ARCHIVEFILEMASK, AGENTID, INBOUNDMAILID, METAFIELD1NAME, METAFIELD1VALUE, RECORDHANDLEFLAG) SELECT 'PRSC_[AGENTID]_[METAFIELD1VALUE]_[MAILID]',@AGENTID,@INBOUNDMAILID,@METAFIELD1NAME,@METAFIELD1VALUE,1";

                        sqlcmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("AGENTID", archiveRequestSettings["AGENTID"]));
                        sqlcmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("INBOUNDMAILID", archiveRequestSettings["INBOUNDMAILID"]));
                        sqlcmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("METAFIELD1NAME", archiveRequestSettings["METAFIELD1NAME"]));
                        sqlcmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("METAFIELD1VALUE", archiveRequestSettings["METAFIELD1VALUE"]));
                        sqlcmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        if (ex != null)
                        {
                            validationError = "FAILED";
                            //LogMessage("info" + error == "" ? "" : "-err", guid + "LogInfo()", "url:" + url, "postcontent:" + postcontent, "response:" + responsecontent, error == "" ? "" : error);
                            LogMessage("err", "db-error", ex.Message);
                        }
                    }
                    finally
                    {
                        try
                        {
                            sqlcmd.Dispose();
                        }
                        catch (Exception exs)
                        {

                        }
                    }
                }
                else
                {
                    validationError = "FAILED";
                    LogMessage("info" + validationError == "" ? "" : "-err", validationError);
                }
            } else
            {
                validationError = validationError.Substring(0, validationError.Length - 1);
                LogMessage("error", validationError);
            }
            return validationError;
        }

        private static System.IO.StreamWriter logstrm = null;

        public static void LogMessage(string msgtype, params string[] args)
        {
            var logfilename = System.AppDomain.CurrentDomain.BaseDirectory + "presenceMailArchive." + DateTime.Now.ToString("yyyy-MM-dd") + ".log";
            if (!System.IO.File.Exists(logfilename))
            {
                if (logstrm != null)
                {
                    logstrm.Flush();
                    logstrm.Close();
                }
                if (logstrm == null)
                {
                    try
                    {
                        logstrm = new System.IO.StreamWriter(logfilename, true);
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
            else
            {
                if (logstrm == null)
                {
                    try
                    {
                        logstrm = new System.IO.StreamWriter(logfilename, true);
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
            if (logstrm != null)
            {
                logstrm.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " " + msgtype + string.Join(" ", args));
                logstrm.Flush();
            }
        }
    }
}
