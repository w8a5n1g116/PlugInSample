using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using PlanSource.Model;
using ScheduleCore;
using Braincase.GanttChart;
using System.IO;
using OfficeOpenXml;
using PlugInSample.Model;
using System.Windows.Forms.DataVisualization.Charting;


/* OrBit的C#插件示例
 * 采用VS-2008-SP1编写
 * 版本号V10.40
 * 本插件适用于OrBit-Browser V10.20　或以上的版本
 * 提供了各种必要的接口(属性、函数(方法\过程))，开发者可以在此基础上进行改编或完善
 * 由:深圳OrBit Systems Inc. OrBit Team提供
 * 发布日期 2009年1月1日
 * 最后修改 2011年10月23日
 */


namespace SchedulePlugInSample
{
    public partial class ScheduleUserControl : UserControl
    {  //类初始化

        /// <summary>
        /// 以下定义排产逻辑对象
        /// </summary>
        SchedulePlan SchedulePlan = null;
        private List<string> MoIdListLocal = new List<string>();

        public delegate void OutSideFunction();
        public event OutSideFunction outsideInvoke;

        public delegate void ComputerDelegate();
        public ComputerDelegate computerDelegate;

        public delegate void RefreshDelegate();
        public RefreshDelegate refreshDelegate;

        public delegate void ErrorDelegate(string message);
        public ErrorDelegate errorDelegate;

        public delegate void GanttDelegate();
        public GanttDelegate ganttDelegate;

        public Venture currentVenture = null;

        public int MoItemGridPageNumber = 0;
        public int MoGridPageNumber = 0;

        /// <summary>
        /// MOList缓存
        /// </summary>
        private List<MoModel> MoListTemp = new List<MoModel>();

        private List<MoItemStepModel> MoStepListTemp = new List<MoItemStepModel>();

        private List<MoItemStepModel> MoStepListCheckTemp = new List<MoItemStepModel>();

        #region ///以下定义OrBit插件与浏览器进行交互的接口变量

        //基本接口变量
        public string PlugInCommand;  //插件的命令ID
        public string PlugInName; //插件的名字(能随语言ID改变)
        public string LanguageId; //语言ID(0-英,1-简,2-繁...8)
        public string ParameterString; //参数字串
        public string RightString; //权限字串(1-Z)
        public string OrBitUserId; //OrBit用户ID
        public string OrBitUserName; //用户名
        public string ApplicationName; //应用系统
        public string ResourceId; //资源(电脑)ID
        public string ResourceName; //资源名
        public bool IsExitQuery; //在插件窗体退出是询问是否要退出，用于提醒数据状态已改变。
        public string UserTicket; //经浏览器认证后的加密票，调用某些WCF服务时需要使用

        //单独调试数据库应用时需用到参数
        public string DatabaseServer; //数据库服务器
        public string DatabaseName;//数据库名
        public string DatabaseUser;//数据库用户
        public string DatabasePassword; //密码

        //各服务器的地址
        public string WcfServerUrl; // WCF或Webservice服务的路径
        public string DocumentServerURL ; //文档服务器URL
        public string PluginServerURL ;//插件服务器URL
        public string RptReportServerURL ; //水晶报表服务器URL

        //回传给浏览器的元对象信息
        public string MetaDataName = "No metadata"; //元对象名
        public string MetaDataVersion = "0.00"; //元对象版本
        public string UIMappingEngineVersion = "0.00"; //UIMapping版本号

        //事件日志类型枚举--1.普通事件，2警告，3错误，4严重错误 ,5表字段变更 ,6其它
        public enum EventLogType { Normal = 1, Warning = 2, Error = 3, FatalError = 4, TableChanged = 5, Other = 6 };

        #endregion

        #region ////以下定义OrBit插件与浏览器交互的接口函数

        /// <summary>
        /// 朗读文本
        /// </summary>
        /// <param name="text">语音内容</param>

        public void SpeechText(string text)
        { //朗读文本
            if (text.Trim() == "") return;
            try
            {
                Type type = this.ParentForm.GetType();
                //调用没有返回值的方法
                Object obj = this.ParentForm;
                type.InvokeMember("SpeechText", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance, null, obj, new object[] { text });
            }
            catch
            {

            }
        }

        /// <summary>
        /// 接收Winsock传过来的消息指令
        /// </summary>
        /// <param name="commandString">命令字串</param>
        public void WinsockMessageReceive(string msgString)
        {
           // this.ParentForm.Activate(); 
            MessageBox.Show(msgString);  
        }

        /// <summary>
        /// 执行浏览器的命令
        /// </summary>
        /// <param name="command">命令ID</param>
        public void RunCommand(string command) 
        {
            try
            {
                Type type = this.ParentForm.GetType();
                //调用没有返回值的方法
                Object obj = this.ParentForm;
                type.InvokeMember("RunCommand", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance, null, obj, new object[] { command });
            }
            catch 
            { 
            
            }
        }
        /// <summary>
        /// 将键值写入OrBit的注册表
        /// </summary>
        /// <param name="ApplicationName">应用系统名</param>
        /// <param name="KeyName">键名</param>
        /// <param name="KeyValue">键值</param>
        public void WriteRegistryKey(string ApplicationName, string KeyName, string KeyValue)
        {
            try
            {
                Type type = this.ParentForm.GetType();
                //调用没有返回值的方法
                Object obj = this.ParentForm;
                type.InvokeMember("GetSetRegistryInfomation", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance, null, obj, new object[] { ApplicationName + "." + KeyName, KeyValue, false });
            }
            catch
            {

            }
        }
        /// <summary>
        /// 从OrBit的注册表读取键值
        /// </summary>
        /// <param name="ApplicationName"></param>
        /// <param name="KeyName"></param>
        /// <returns></returns>
        public string ReadRegistryKey(string ApplicationName, string KeyName)
        {  //读取注册表

            try
            {
                Type type = this.ParentForm.GetType();
                //调用没有返回值的方法
                Object obj = this.ParentForm;

                return (string)type.InvokeMember("GetSetRegistryInfomation", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance, null, obj, new object[] { ApplicationName + "." + KeyName, "", true });
            }
            catch
            {
                return "";
            }

        }

        /// <summary>
        /// 采用OrBit-MES的WCF服务器上传文件
        public string UploadFile(string Filename, string SafeFileName, string UploadFilePath)
        {  

            try
            {
                Type type = this.ParentForm.GetType();
                //调用没有返回值的方法
                Object obj = this.ParentForm;

                return (string)type.InvokeMember("UploadFile", BindingFlags.InvokeMethod | BindingFlags.Public |
                    BindingFlags.Instance, null, obj, new object[] { Filename, SafeFileName, UploadFilePath });
            }
            catch
            {
                return "";
            }

        }

        /// <summary>
        /// 在内置浏览器中打开一个WebURL
        /// </summary>
        /// <param name="URL"></param>
         public void OpenWebUrl(string URL) 
        {
            try
            {
                //用反射的方法直接调用浏览器来浏览文件                
                Type type = this.ParentForm.GetType();
                Object obj = this.ParentForm;
                type.InvokeMember("OpenIEURL", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance, null, obj, new object[] { URL });
            }
            catch 
            { 
            
            }
        }

        /// <summary>
        /// 将信息送往浏览器状态条
        /// </summary>
        /// <param name="Message">信息</param>
        public void SendToStatusBar(string Message)
        {
            try
            {
                Type type = this.ParentForm.GetType();
                //调用没有返回值的方法
                Object obj = this.ParentForm;
                type.InvokeMember("SendToStatusBar", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance, null, obj, new object[] { Message });
            }
            catch
            {

            }
        }

        /// <summary>
        /// 更改浏览器进度条
        /// </summary>
        /// <param name="Maximum">最大值</param>
        /// <param name="Value">当前值</param>        
        public void ChangeProgressBar(int Maximum, int Value)
        {
            try
            {
                Type type = this.ParentForm.GetType();
                //调用没有返回值的方法
                Object obj = this.ParentForm;
                type.InvokeMember("ChangeProgressBar", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance, null, obj, new object[] { Maximum, Value });
            }
            catch
            {

            }
        }

        /// <summary>
        /// 写入事件日志-重载1
        /// </summary>
        /// <param name="eventLogType">1.普通事件，2警告，3错误，4严重错误 ,5表字段变更 ,6其它</param>
        /// <param name="eventLog">事件具体内容</param>
        public void WriteToEventLog(string eventLogType, string eventLog)
        {
            if (eventLogType == "1") WriteToEventLog(EventLogType.Normal, eventLog);
            if (eventLogType == "2") WriteToEventLog(EventLogType.Warning, eventLog);
            if (eventLogType == "3") WriteToEventLog(EventLogType.Error, eventLog);
            if (eventLogType == "4") WriteToEventLog(EventLogType.FatalError, eventLog);
            if (eventLogType == "5") WriteToEventLog(EventLogType.TableChanged, eventLog);
            if (eventLogType == "6") WriteToEventLog(EventLogType.Other, eventLog);
        }
        /// <summary>
        /// 写入事件日志-原型
        /// </summary>
        /// <param name="eventLogType">1.普通事件，2警告，3错误，4严重错误 ,5表字段变更 ,6其它</param>
        /// <param name="eventLog">事件具体内容</param>
        public void WriteToEventLog(EventLogType eventLogType, string eventLog)
        {
            try
            {
                Type type = this.ParentForm.GetType();
                //调用没有返回值的方法
                Object obj = this.ParentForm;
                type.InvokeMember("WriteToEventLog", BindingFlags.InvokeMethod | BindingFlags.Public |
                                BindingFlags.Instance, null, obj, new object[] { (int)eventLogType,this.ParentForm.Text,eventLog });
            }
            catch
            {

            }
        }

        //更新一个由客户端传回的记录集
        public bool UpdateDataSetBySQL(DataSet DataSetWithSQL, string SQLString)
        {
           try{
            if (this.Parent.Name.ToString() != "FormPlugIn")
                {　//在插件调试环境下运行时，用ADO.NET直连 
                    string ConnectionString = "Data Source=" + DatabaseServer +
                            ";Initial Catalog=" + DatabaseName +
                            ";password=" + DatabasePassword +
                            ";Persist Security Info=True;User ID=" + DatabaseUser;
                    using (SqlConnection conn = new SqlConnection(ConnectionString))
                    {
                       
                        conn.ConnectionString = ConnectionString;
                        conn.Open();
                        SqlDataAdapter da = new SqlDataAdapter(SQLString, conn);
                        SqlCommandBuilder objCommandBuilder = new SqlCommandBuilder(da);
                        da.Update(DataSetWithSQL.Tables[0]);
                        conn.Close();
                        return true;
                    }
                }
                else
                {  //在浏览器下运行时，直接调用浏览器的RunStoredProcedure,

                    Type type = this.ParentForm.GetType();
                    Object obj = this.ParentForm;
                    bool resultR = true;
                    Object[] myArgs = new object[] { DataSetWithSQL, SQLString };
                    resultR = (bool)type.InvokeMember("UpdateDataSetBySQL", BindingFlags.InvokeMethod | BindingFlags.Public |
                                    BindingFlags.Instance, null, obj, myArgs);
                    return resultR;

                }
            }
            catch (Exception er)
            {
                MessageBox.Show(er.ToString(), this.ParentForm.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        //创建一个用于执行存储过程的结构
        public DataSet CreateIOParameterDataSet() 
        {
            DataSet ds = new DataSet();//声明一个DataSet对象
            ds.Tables.Add(new DataTable()); //添加一个表
            ds.Tables[0].Columns.Add(new DataColumn("FieldName"));
            ds.Tables[0].Columns.Add(new DataColumn("FieldType"));
            ds.Tables[0].Columns.Add(new DataColumn("IsOutput"));
            ds.Tables[0].Columns.Add(new DataColumn("FieldValue"));

            return ds;        
        }

        //创建执行存储过程的结构记录集的输入输出参数
        public void CreateIOParameter(DataSet IOParameterDataSet,string FieldName,string FieldType,bool IsOutput, string FieldValue)
        {
         
            if (FieldType.ToUpper()!="Bit".ToUpper() &&
                FieldType.ToUpper()!="DateTime".ToUpper() &&
                FieldType.ToUpper()!="Decimal".ToUpper() &&
                FieldType.ToUpper()!="Int".ToUpper() &&
                FieldType.ToUpper()!="NVarChar".ToUpper())
                FieldType ="NVarChar";

            bool isUpdate=false;
            for (int i = 0; i < IOParameterDataSet.Tables[0].Rows.Count; i++) 
            {
                if (IOParameterDataSet.Tables[0].Rows[i]["FieldName"].ToString().Trim().ToUpper() == FieldName.Trim().ToUpper() )
                {
                    IOParameterDataSet.Tables[0].Rows[i]["FieldType"] = FieldType.ToUpper();
                    IOParameterDataSet.Tables[0].Rows[i]["IsOutput"] = IsOutput;
                    IOParameterDataSet.Tables[0].Rows[i]["FieldValue"] = FieldValue;
                    isUpdate=true;
                    break ;
                }            
            
            }
            if (isUpdate == false) 
            {
                DataRow dr = IOParameterDataSet.Tables[0].NewRow();//添加一个新行对象
                dr["FieldName"] = FieldName;
                dr["FieldType"] = FieldType;
                dr["IsOutput"] = IsOutput;
                dr["FieldValue"] = FieldValue;
                IOParameterDataSet.Tables[0].Rows.Add(dr);  //将新添加的行添加到表中            
            }
            IOParameterDataSet.AcceptChanges();            
        }

        //取得输入输出参数的值
        public string GetIOParameterValue(DataSet IOParameterDataSet, string FieldName)
        {   
            for (int i = 0; i < IOParameterDataSet.Tables[0].Rows.Count; i++)
            {
                if (IOParameterDataSet.Tables[0].Rows[i]["FieldName"].ToString().Trim().ToUpper() == FieldName.Trim().ToUpper())
                {
                    return IOParameterDataSet.Tables[0].Rows[i]["FieldValue"].ToString();                              
                }

            }
            return "";
        }

        //更新一个由客户端传回的记录集
        public DataSet ExecSP(string SpName, DataSet IOParameterDataSet, out int Return_value, out  DataSet IOParameterDataSetReturn)
        {
            Return_value = -1;
            IOParameterDataSetReturn = null;
            try
            {
                if (this.Parent.Name.ToString() != "FormPlugIn")
                {　//在插件调试环境下运行时，用ADO.NET直连 
                   DataSet ds= execSPInner(SpName, IOParameterDataSet, out Return_value) ;
                   IOParameterDataSetReturn = IOParameterDataSet;
                   return ds;
                }
                else
                {  //在浏览器下运行时，直接调用浏览器的RunStoredProcedure,
                   
                    Type type = this.ParentForm.GetType();
                    Object obj = this.ParentForm;
                    DataSet resultR = new DataSet() ;                     
                    Object[] myArgs = new object[] { SpName, IOParameterDataSet, Return_value, IOParameterDataSetReturn };
                    resultR = (DataSet)type.InvokeMember("ExecSP", BindingFlags.InvokeMethod | BindingFlags.Public |
                                    BindingFlags.Instance, null, obj, myArgs);
                    IOParameterDataSetReturn = (DataSet)myArgs[3];
                    Return_value = (int)myArgs[2];                    
                    return resultR;

                }
            }
            catch (Exception er)
            {
                MessageBox.Show(er.ToString(), this.ParentForm.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }


        //纯DataSet版本的存储过程执行接口
        private DataSet execSPInner(string SpName, DataSet IOParameterDataSet, out int Return_value) 
        {
            SqlCommand cmdDs = new SqlCommand();
            DataSet ReturnDataSet = new DataSet();
            Return_value = 0;
            //将DataSet转化为SqlCommand
                if (IOParameterDataSet != null && IOParameterDataSet.Tables.Count > 0 && IOParameterDataSet.Tables[0].Columns.Count > 0)
                {
                    cmdDs.CommandText = SpName;
                    for (int i = 0; i < IOParameterDataSet.Tables[0].Rows.Count; i++)
                    {
                        bool isOutput = false;
                        string paraName = IOParameterDataSet.Tables[0].Rows[i]["FieldName"].ToString();
                        
                        if (IOParameterDataSet.Tables[0].Rows[i]["IsOutput"].ToString().ToUpper().Trim()=="TRUE")
                        {
                            isOutput = true;                           
                        }

                        string dataType = IOParameterDataSet.Tables[0].Rows[i]["FieldType"].ToString();

                        if (dataType.ToUpper() == "Bit".ToUpper())
                            cmdDs.Parameters.Add(new SqlParameter(paraName, SqlDbType.Bit));
                        if (dataType.ToUpper() == "DateTime".ToUpper())
                            cmdDs.Parameters.Add(new SqlParameter(paraName, SqlDbType.DateTime));
                        if (dataType.ToUpper() == "Decimal".ToUpper())
                            cmdDs.Parameters.Add(new SqlParameter(paraName, SqlDbType.Decimal));
                        if (dataType.ToUpper() == "Int".ToUpper())
                            cmdDs.Parameters.Add(new SqlParameter(paraName, SqlDbType.Int));
                        if (dataType.ToUpper() == "NVarChar".ToUpper())
                            cmdDs.Parameters.Add(new SqlParameter( paraName, SqlDbType.NVarChar, 1000));


                        cmdDs.Parameters[paraName].Value = IOParameterDataSet.Tables[0].Rows[i]["FieldValue"];
                        if (isOutput == true) cmdDs.Parameters[paraName].Direction = ParameterDirection.Output;
                        

                    }

                    cmdDs.Parameters.Add("RETURN_VALUE", SqlDbType.Int);
                    cmdDs.Parameters["RETURN_VALUE"].Direction = ParameterDirection.ReturnValue;


                    // 执行SqlCommand

                    string ConnectionString = "Data Source=" + DatabaseServer +
                        ";Initial Catalog=" + DatabaseName +
                        ";password=" + DatabasePassword +
                        ";Persist Security Info=True;User ID=" + DatabaseUser;

                    SqlConnection conn = new SqlConnection();
                    conn.ConnectionString = ConnectionString;
                    conn.Open();

                    cmdDs.CommandTimeout = conn.ConnectionTimeout;
                    cmdDs.Connection = conn;
                    cmdDs.CommandType = CommandType.StoredProcedure;


                    DataSet ds = new DataSet("WCFSQLDataSet");
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = cmdDs;
                    adapter.Fill(ds, "WCFSQLDataSet");
                    ReturnDataSet = ds;
                    conn.Close();

                    //将返回值类型的参数送回到DataSet
                    for (int i = 0; i < cmdDs.Parameters.Count; i++)
                    {
                        if (cmdDs.Parameters[i].Direction.ToString() == "Output")
                        {
                            for (int ii = 0; ii < IOParameterDataSet.Tables[0].Rows.Count; ii++)
                            {
                            
                                if (IOParameterDataSet.Tables[0].Rows[ii]["IsOutput"].ToString().ToUpper().Trim() == "TRUE" &&
                                    IOParameterDataSet.Tables[0].Rows[ii]["FieldName"].ToString().ToUpper() == cmdDs.Parameters[i].ToString().ToUpper())
                                {
                                    IOParameterDataSet.Tables[0].Rows[ii]["FieldValue"] = cmdDs.Parameters[i].Value;
                                }
                            }

                        }

                        if (cmdDs.Parameters[i].Direction.ToString() == "ReturnValue")
                        {
                            Return_value = (int)cmdDs.Parameters["RETURN_VALUE"].Value;
                        }
                    }
                    return ReturnDataSet;
                }
                
                return null;
         }


        /// <summary>
        /// 通用执行存储过程程序.
        /// </summary>
        /// <param name="SQLCmd">传入的SqlCommand对象</param>
        /// <param name="ReturnDataSet">执行存储过程后返回的数据集</param>
        /// <param name="ReturnValue">执行存储过程的返回值</param>
        /// <returns>将SQLCmd执行后的参数刷新并传回，主要返回存储过程中的out参数</returns>
        public SqlCommand RunStoredProcedure(SqlCommand SQLCmd, out DataSet ReturnDataSet, out int ReturnValue) 
        {
            ReturnValue = 0;
            try
            {
                if (this.Parent.Name.ToString() != "FormPlugIn")
                {　//在插件调试环境下运行时，用ADO.NET直连 
                    string ConnectionString = "Data Source=" + DatabaseServer +
                            ";Initial Catalog=" + DatabaseName +
                            ";password=" + DatabasePassword +
                            ";Persist Security Info=True;User ID=" + DatabaseUser;
                    using (SqlConnection conn = new SqlConnection(ConnectionString))
                    {
                        conn.Open();
                        SQLCmd.Connection = conn;
                        SQLCmd.CommandType = CommandType.StoredProcedure;
                        SQLCmd.CommandTimeout = conn.ConnectionTimeout;
                        SQLCmd.Parameters.Add("RETURN_VALUE", SqlDbType.Int);
                        SQLCmd.Parameters["RETURN_VALUE"].Direction = ParameterDirection.ReturnValue;

                        SqlDataAdapter adapter = new SqlDataAdapter();
                        adapter.SelectCommand = SQLCmd;

                        DataSet ds = new DataSet("WCFSQLDataSet");
                        adapter.Fill(ds, "WCFSQLDataSet");

                        ReturnDataSet = ds;
                        conn.Close();
                        ReturnValue = (int)SQLCmd.Parameters["RETURN_VALUE"].Value;                        
                        return SQLCmd;
                    }
                }
                else
                {  //在浏览器下运行时，直接调用浏览器的RunStoredProcedure,

                    Type type = this.ParentForm.GetType();
                    Object obj = this.ParentForm;

                    SqlCommand cmd = new SqlCommand();
                    DataSet rds = new DataSet();

                    ReturnDataSet = null;

                    Object[] myArgs = new object[] { SQLCmd, ReturnDataSet, ReturnValue };
                    cmd = (SqlCommand)type.InvokeMember("RunStoredProcedure", BindingFlags.InvokeMethod | BindingFlags.Public |
                                    BindingFlags.Instance, null, obj, myArgs);
                    ReturnValue = (int)myArgs[2];
                    ReturnDataSet = (DataSet)myArgs[1];
                    return cmd;

                }
            }
            catch (Exception er)
            {
                ReturnDataSet = null;
                MessageBox.Show(er.ToString(), this.ParentForm.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

        }

        /// <summary>
        /// 通用执行存储过程程序，并支侍向存储过程上传数据集,需SQL-Server 2008对表值类型传入参数的支持
        /// </summary>
        /// <param name="SQLCmd">传入的SqlCommand对象</param>
        /// <param name="UserDataSet">用于上传数据的UserDataSet</param>
        /// <param name="ReturnDataSet">执行存储过程后返回的数据集</param>
        /// <param name="ReturnValue">执行存储过程的返回值</param>
        /// <returns>将SQLCmd执行后的参数刷新并传回，主要返回存储过程中的out参数</returns>
        public SqlCommand RunSPUploadDataSet(SqlCommand SQLCmd, DataSet UserDataSet, out DataSet ReturnDataSet,out int ReturnValue)
        {
            ReturnValue = 0;
            try
            {
                if (this.Parent.Name.ToString() != "FormPlugIn")
                {　//在插件调试环境下运行时，用ADO.NET直连 
                    string ConnectionString = "Data Source=" + DatabaseServer +
                            ";Initial Catalog=" + DatabaseName +
                            ";password=" + DatabasePassword +
                            ";Persist Security Info=True;User ID=" + DatabaseUser;
                    using (SqlConnection conn = new SqlConnection(ConnectionString))
                    {
                        conn.Open();
                        SQLCmd.Connection = conn;
                        SQLCmd.CommandType = CommandType.StoredProcedure;
                        SQLCmd.CommandTimeout = conn.ConnectionTimeout;
                        SQLCmd.Parameters.Add("RETURN_VALUE", SqlDbType.Int);
                        SQLCmd.Parameters["RETURN_VALUE"].Direction = ParameterDirection.ReturnValue;

                        //处理一次性从客户端传过来的DataSet，并将它的Table名作为参数名使用
                        if (UserDataSet != null && UserDataSet.Tables.Count > 0)
                        {
                            for (int k = 0; k < UserDataSet.Tables.Count; k++)
                            {
                                string tableName = UserDataSet.Tables[k].TableName.ToString().Trim();
                                if (tableName != "")
                                {
                                    SQLCmd.Parameters.AddWithValue("@" + tableName, UserDataSet.Tables[k]);
                                }
                            }
                        }

                        SqlDataAdapter adapter = new SqlDataAdapter();
                        adapter.SelectCommand = SQLCmd;

                        DataSet ds = new DataSet("WCFSQLDataSet");
                        adapter.Fill(ds, "WCFSQLDataSet");

                        ReturnDataSet = ds;
                        conn.Close();
                        ReturnValue = (int)SQLCmd.Parameters["RETURN_VALUE"].Value;
                        return SQLCmd;
                    }
                }
                else
                {  //在浏览器下运行时，直接调用浏览器的RunStoredProcedure,
                                        
                    Type type = this.ParentForm.GetType();
                    Object obj = this.ParentForm;

                    SqlCommand cmd = new SqlCommand();
                    DataSet rds = new DataSet();

                    ReturnDataSet = null;

                    Object[] myArgs = new object[] { SQLCmd, UserDataSet, ReturnDataSet, ReturnValue };
                    cmd = (SqlCommand)type.InvokeMember("RunSPUploadDataSet", BindingFlags.InvokeMethod | BindingFlags.Public |
                                    BindingFlags.Instance, null, obj, myArgs);
                    ReturnValue = (int)myArgs[3];
                    ReturnDataSet = (DataSet)myArgs[2];
                    return cmd;

                }
            }
            catch (Exception er)
            {
                ReturnDataSet = null;
                MessageBox.Show(er.ToString(), this.ParentForm.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

        }

        /// <summary>
        /// //为插件提供可以任意一个WCF的地址上直接调用任意存储过程的接口,并支持上传数据集UserDataSet。
        /// </summary>
        /// <param name="WCFUrl">一个任意的WCFUrl地址，需要完整的写出*.svc</param>
        /// <param name="SQLCmd">传入的SqlCommand对象</param>
        /// <param name="UserDataSet">用于上传数据的UserDataSet</param>
        /// <param name="ReturnDataSet">执行存储过程后返回的数据集</param>
        /// <param name="ReturnValue">执行存储过程的返回值</param>
        /// <returns>将SQLCmd执行后的参数刷新并传回，主要返回存储过程中的out参数</returns>
        public SqlCommand RunSPUploadDataSetByWCFUrl(string WCFUrl, SqlCommand SQLCmd, DataSet UserDataSet, out DataSet ReturnDataSet, out int ReturnValue)
        {
            ReturnValue = 0;
            try
            {
                if (this.Parent.Name.ToString() != "FormPlugIn")
                {　//在插件调试环境下运行时，用ADO.NET直连 
                    MessageBox.Show("本函数必须在OrBit浏览器环境下执行！"); 
                    ReturnDataSet = null;
                    return null;
                    
                }
                else
                {  //在浏览器下运行时，直接调用浏览器的RunStoredProcedure,

                    Type type = this.ParentForm.GetType();
                    Object obj = this.ParentForm;

                    SqlCommand cmd = new SqlCommand();
                    DataSet rds = new DataSet();

                    ReturnDataSet = null;

                    Object[] myArgs = new object[] {WCFUrl, SQLCmd, UserDataSet, ReturnDataSet, ReturnValue };                    
                    cmd = (SqlCommand)type.InvokeMember("RunSPUploadDataSetByWCFUrl", BindingFlags.InvokeMethod | BindingFlags.Public |
                                    BindingFlags.Instance, null, obj, myArgs);
                    ReturnValue = (int)myArgs[4];
                    ReturnDataSet = (DataSet)myArgs[3];
                    return cmd;

                }
            }
            catch (Exception er)
            {
                ReturnDataSet = null;
                MessageBox.Show(er.ToString(), this.ParentForm.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

        }

        /// <summary>
        /// 执行一个指定的SQL字串，并返回一个记录集
        /// 在浏览器下执行时，直接调用浏览器的WCF服务器来传送记录集
        /// </summary>
        /// <param name="SQLString">SQL字串</param>
        /// <returns>返回的记录集</returns>
        public DataSet GetDataSetWithSQLString(string SQLString)
        {
            try
            {
                if (this.Parent.Name.ToString() != "FormPlugIn")
                {　//在插件调试环境下运行时，用ADO.NET直连 
                    string ConnectionString = "Data Source=" + DatabaseServer +
                                        ";Initial Catalog=" + DatabaseName +
                                        ";password=" + DatabasePassword +
                                        ";Persist Security Info=True;User ID=" + DatabaseUser; 
                    using (SqlConnection conn = new SqlConnection(ConnectionString))
                    {
                        conn.Open();
                        SqlCommand comm = new SqlCommand();
                        comm.Connection = conn;
                        comm.CommandText = SQLString;

                        comm.CommandType = CommandType.Text ;
                        comm.CommandTimeout = conn.ConnectionTimeout;

                        DataSet ds = new DataSet("SQLDataSet");
                        SqlDataAdapter adapter = new SqlDataAdapter();
                        adapter.SelectCommand = comm;
                        adapter.Fill(ds, "SQLDataSet");
                       
                        conn.Close();
                        return ds;
                    }
                }
                else
                {　//在浏览器下运行时，直接调用浏览器的GetDataSetWithSQLString,
                    //通过WCF服务器返回记录集

                    Type type = this.ParentForm.GetType();
                    Object obj = this.ParentForm;
                    DataSet ds = new DataSet("SQLDataSet");
                    ds = (DataSet)type.InvokeMember("GetDataSetWithSQLString", BindingFlags.InvokeMethod | BindingFlags.Public |
                                    BindingFlags.Instance, null, obj, new object[] { SQLString });
                    return ds;

                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            } 
        }

        /// <summary>
        /// 执行一个指定的SQL字串，并从OLAP服务器上返回一个记录集
        /// 在浏览器下执行时，直接调用浏览器的WCF服务器来传送记录集
        /// </summary>
        /// <param name="SQLString">SQL字串</param>
        /// <returns>返回的记录集</returns>
        public DataSet GetDataSetWithSQLStringFromOLAP(string SQLString)
        {
            try
            {
                if (this.Parent.Name.ToString() != "FormPlugIn")
                {　//在插件调试环境下运行时，用ADO.NET直连 
                    string ConnectionString = "Data Source=" + DatabaseServer +
                                        ";Initial Catalog=" + DatabaseName +
                                        ";password=" + DatabasePassword +
                                        ";Persist Security Info=True;User ID=" + DatabaseUser;
                    using (SqlConnection conn = new SqlConnection(ConnectionString))
                    {
                        conn.Open();
                        SqlCommand comm = new SqlCommand();
                        comm.Connection = conn;
                        comm.CommandText = SQLString;

                        comm.CommandType = CommandType.Text;
                        comm.CommandTimeout = conn.ConnectionTimeout;

                        DataSet ds = new DataSet("SQLDataSet");
                        SqlDataAdapter adapter = new SqlDataAdapter();
                        adapter.SelectCommand = comm;
                        adapter.Fill(ds, "SQLDataSet");

                        conn.Close();
                        return ds;
                    }
                }
                else
                {　//在浏览器下运行时，直接调用浏览器的GetDataSetWithSQLString,
                    //通过WCF服务器返回记录集

                    Type type = this.ParentForm.GetType();
                    Object obj = this.ParentForm;
                    DataSet ds = new DataSet("SQLDataSet");
                    ds = (DataSet)type.InvokeMember("GetDataSetWithSQLStringFromOLAP", BindingFlags.InvokeMethod | BindingFlags.Public |
                                    BindingFlags.Instance, null, obj, new object[] { SQLString });
                    return ds;

                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), this.ParentForm.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        #region // GetUIText为浏览器提供的语言切换函数，可以用它返回控件或提示信息不同语种的内容。
        // GetUIText提供了几种形式的重载，以方便使用

        /// <summary>
        /// 返回控件或提示信息不同语种的内容-重载4
        /// </summary>
        /// <param name="module">模块名</param>
        /// <returns></returns>
        public String GetUIText(string module)
        {
            string s= GetUIText("[Public Text]", module, "", "").Trim() ;
            if (s == string.Empty)
            {
                return module;
            }
            else 
            {
                return s;
            }

        }
        /// <summary>
        /// 返回控件或提示信息不同语种的内容-重载3
        /// </summary>
        /// <param name="module">模块名</param>
        /// <param name="defaultText">默认文本,必须是英文</param>
        /// <returns></returns>
        public string GetUIText(string module, string defaultText)
        {
            string s= GetUIText("[Public Text]", module, "", defaultText).Trim() ;
            if (s == string.Empty)
            {
                return defaultText;
            }
            else 
            {
                return s;
            }
        }
        /// <summary>
        /// 返回控件或提示信息不同语种的内容-重载2
        /// </summary>
        /// <param name="module">模块名</param>
        /// <param name="control">控件对象</param>
        /// <returns></returns>
        public string GetUIText(string module, Control control)
        {
            string s = GetUIText(this.ParentForm.Text, module, control.Name, control.Text);
            if (s == string.Empty)
            {
                return control.Text;
            }
            else
            {
                return s;
            }

        }
        /// <summary>
        /// 返回控件或提示信息不同语种的内容-重载1
        /// </summary>
        /// <param name="module">模块名</param>
        /// <param name="controlName">控件名</param>
        /// <param name="defaultText">默认文本,必须是英文</param>
        /// <returns></returns>
        public string GetUIText(string module, string controlName, string defaultText)
        {
            string s = GetUIText(this.ParentForm.Text, controlName, defaultText);
            if (s == string.Empty)
            {
                return defaultText;
            }
            else
            {
                return s;
            }
        }
        /// <summary>
        /// 返回控件或提示信息不同语种的内容-基本型式
        /// </summary>
        /// <param name="owner">所有者</param>
        /// <param name="module">模块名</param>
        /// <param name="controlName">控件名</param>
        /// <param name="defaultText">默认文本,必须是英文</param>
        /// <returns></returns>
        public string GetUIText(string owner,string module, string controlName, string defaultText)
        {
            try
            {
                Type type = this.ParentForm.GetType();
                //调用没有返回值的方法
                Object obj = this.ParentForm;
                string s = (string)type.InvokeMember("GetUIText", BindingFlags.InvokeMethod | BindingFlags.Public |
                                BindingFlags.Instance, null, obj, new object[] { owner, module, controlName, defaultText });
                if (s != null)
                {
                    return s;
                }
                else 
                {
                    return defaultText;
                }
            }
            catch
            {
                return defaultText;
            }

        }
        #endregion

        #endregion

        

        /// <summary>
        /// 插件入口,由.NET自动生成
        /// </summary>
        public ScheduleUserControl()
        {

            InitializeComponent();　//插件控件布局(.NET的默认过程)
            initializeVariable(); //插件变量初始化
            SwitchUI();　　//切换语言
            BindMoGridViewColumn();
            BindMoStepGridViewColumn();
        }

        /// <summary>
        /// 本私有函数对插件各接口变量进行初始化，赋予默认值
        /// 调试环境下这些值不变，通过浏览器执行时，
        /// 这些变量将会根据系统环境被重新赋值。
        /// </summary>
        private void initializeVariable() 
        {
            PlugInCommand = "MYPGN";
            PlugInName = "我的插件";
            LanguageId = "0";  //(0-英,1-简,2-繁...8)
            ParameterString = "";
            RightString = "(0)";

            OrBitUserId = "";
            OrBitUserName = "调试者";
            ApplicationName = "DEBUG";
            ResourceId = "RES0000XXX";
            ResourceName = "YourPC";

            //这里需要根据实际的数据库环境进行改写
            DatabaseServer = "HenryX200";
            DatabaseName = "OrBitXI";
            DatabaseUser = "SA";
            DatabasePassword = "sasasa"; 

            WcfServerUrl = "http://192.168.41.58/OrBitWCFService_Savic";
            DocumentServerURL = ""; //文档服务器URL
            PluginServerURL = "http://192.168.41.58/OrBitPluginService_Savic/";//插件服务器URL
            RptReportServerURL = "http://henryx61/RptExamples/"; //水晶报表服务器URL

            UserTicket = "";
            IsExitQuery = false;
        }

        /// <summary>
        /// 插件加载事件，由.NET自动生成
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UserControl1_Load(object sender, EventArgs e)
        {
            this.Dock = DockStyle.Fill;

            computerDelegate = new ComputerDelegate(computerCallBack);
            refreshDelegate = new RefreshDelegate(refreshCallBack);
            errorDelegate = new ErrorDelegate(errorCallBack);
            ganttDelegate = new GanttDelegate(GanttCallBack);
            OutSideFunction f1 = new OutSideFunction(outSideInvokeCallBack);
            outsideInvoke += f1;

            BuildSchedulePlan();

            GetMoGridViewData();
            GetMoStepGridViewData();

            ShowMoGridViewData();
            ShowMoStepGridViewData();

            List<Venture> venturelist = SchedulePlan.GetVentureByFactory();

            listBox1.DataSource = venturelist;
            listBox1.DisplayMember = "VentureName";

            //InitGanttChart();

            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now.AddDays(7);


        }

        /// <summary>
        /// SwitchUI为OrBit浏览器提供的统一语言切换函数,
        /// 能根据浏览器的语言环境切换界面语言
        /// 如无多语言切换需求，本函数可以置空。
        /// 如有多语言切换需求，请自行完善本函数的内容
        /// 此函数不可改名或删除，否则将导致浏览器无法对其切换语言。
        /// </summary>
        public void SwitchUI()
        {
            
            tsbtnNew.ToolTipText = GetUIText("Create");
            tsbtnModify.ToolTipText = GetUIText("Modify");
            tsbtnSave.ToolTipText = GetUIText("Save");
            tsbtnCancel.ToolTipText = GetUIText("Cancel");
            tsbtnDelete.ToolTipText = GetUIText("Delete");
            tsbtnAprove.ToolTipText = GetUIText("Aprove");
            tsbtnFind.ToolTipText = GetUIText("Find");
            tsbtnPrint.ToolTipText = GetUIText("Print");
            tsbtnImport.ToolTipText = GetUIText("Import data");
            tsbtnExport.ToolTipText = GetUIText("Export date");
            tsbtnExit.ToolTipText = GetUIText("Exit");


        }


        /// <summary>
        /// PluginUnload函数为OrBit浏览器关闭窗体时触发的过程,可以用于示范各种对象
        /// 此函数不可改名或删除
        /// 但里面内容允许修改。
        /// </summary>
        public void PluginUnload()
        {

           
        }

        //演示如何退出插件
        private void tsbtnExit_Click(object sender, EventArgs e)
        {             
            this.ParentForm.Close(); 

        }



        private void tsbtnNew_Click(object sender, EventArgs e)
        {
           


        }

        private static DataTable UDatatable() 
        { 
             DataTable dt = new DataTable("dt"); 
             DataColumn[] dtc = new DataColumn[6];
             dtc[0] = new DataColumn("_LsId", System.Type.GetType("System.String"));
             dtc[1] = new DataColumn("Product", System.Type.GetType("System.String"));
             dtc[2] = new DataColumn("Base", System.Type.GetType("System.String"));
             dtc[3] = new DataColumn("NGReason", System.Type.GetType("System.String"));
             dtc[4] = new DataColumn("Qty", System.Type.GetType("System.String"));
             dtc[5] = new DataColumn("A", System.Type.GetType("System.String"));  
             dt.Columns.AddRange(dtc); return dt; 
        } 
     

        private void tsbtnPrint_Click(object sender, EventArgs e)
        {
            
        }

       
        public void BindMoGridViewColumn()
        {
            

            MoGridView.Columns.Add("MOId", "MOID");
            MoGridView.Columns.Add("MOName", "单据编号");
            MoGridView.Columns.Add("ProductId", "产品编码");
            MoGridView.Columns.Add("ProductName", "产品名称");
            MoGridView.Columns.Add("PlannedDateFrom", "计划开始时间");
            MoGridView.Columns.Add("PlannedDateTo", "计划结束时间");
            MoGridView.Columns.Add("MOStatus", "计划状态");
            MoGridView.Columns.Add("SapDeliveryDate", "订单交期");

            MoGridView.Columns["MOId"].DataPropertyName = "MOId";
            MoGridView.Columns["MOName"].DataPropertyName = "MOName";//"单据编号";
            MoGridView.Columns["ProductId"].DataPropertyName = "ProductId";
            MoGridView.Columns["ProductName"].DataPropertyName = "ProductName";
            MoGridView.Columns["PlannedDateFrom"].DataPropertyName = "PlannedDateFrom";
            MoGridView.Columns["PlannedDateTo"].DataPropertyName = "PlannedDateTo";
            MoGridView.Columns["MOStatus"].DataPropertyName = "MOStatus";
            MoGridView.Columns["SapDeliveryDate"].DataPropertyName = "SapDeliveryDate";

            MoGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect; //允许选中行
            MoGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            MoGridView.ReadOnly = true;
            MoGridView.AllowUserToAddRows = false;

            foreach (DataGridViewColumn column in MoGridView.Columns)
            {
                //设置自动排序
                column.SortMode = DataGridViewColumnSortMode.Programmatic;
            }

            MoGridView.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable;

        }

        public void GetMoGridViewData()
        {
            MoListTemp = SchedulePlan.GetMoList(MoGridPageNumber);
        }

        public void ShowMoGridViewData()
        {
            List<MoModel> moList = MoListTemp;

            MoGridView.AutoGenerateColumns = false;
            MoBindingSource.DataSource = moList;
            MoGridView.DataSource = MoBindingSource;

            MoGridView.Refresh();
        }

        public void BindMoStepGridViewColumn()
        {
            //MoStepGridView.Columns.Add("MOId", "MOID");
            //MoStepGridView.Columns.Add("MOItemTaskStepId", "工序计划ID");
            MoStepGridView.Columns.Add("MOName", "单据编号");
            MoStepGridView.Columns.Add("ProductId", "产品编码");
            MoStepGridView.Columns.Add("ProductName", "产品名称");
            MoStepGridView.Columns.Add("Specification", "工序");
            MoStepGridView.Columns.Add("Venture", "经营体");
            MoStepGridView.Columns.Add("PlannedDateFrom", "计划开始时间");
            MoStepGridView.Columns.Add("PlannedDateTo", "计划结束时间");
            MoStepGridView.Columns.Add("TaskStatus", "工序计划状态");
            MoStepGridView.Columns.Add("SapDeliveryDate", "订单交期");

            //MoStepGridView.Columns["MOId"].DataPropertyName = "MOId";
            //MoStepGridView.Columns["MOItemTaskStepId"].DataPropertyName = "MOItemTaskStepId";
            MoStepGridView.Columns["MOName"].DataPropertyName = "MOName";
            MoStepGridView.Columns["ProductId"].DataPropertyName = "ProductId";
            MoStepGridView.Columns["ProductName"].DataPropertyName = "ProductName";
            MoStepGridView.Columns["Specification"].DataPropertyName = "Specification";
            MoStepGridView.Columns["Venture"].DataPropertyName = "Venture";
            MoStepGridView.Columns["PlannedDateFrom"].DataPropertyName = "PlannedDateFrom";
            MoStepGridView.Columns["PlannedDateTo"].DataPropertyName = "PlannedDateTo";
            MoStepGridView.Columns["TaskStatus"].DataPropertyName = "TaskStatus";
            MoStepGridView.Columns["SapDeliveryDate"].DataPropertyName = "SapDeliveryDate";

            MoStepGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect; //允许选中行
            MoStepGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            MoStepGridView.ReadOnly = true;
            MoStepGridView.AllowUserToAddRows = false;

            foreach (DataGridViewColumn column in MoStepGridView.Columns)
            {
                //设置自动排序
                column.SortMode = DataGridViewColumnSortMode.Programmatic;
            }

            MoStepGridView.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable;
        }

        public void GetMoStepGridViewData()
        {
            MoStepListTemp = SchedulePlan.GetScheduledMoItemStepModelList(0);
        }

        public void ShowMoStepGridViewData()
        {

            List<MoItemStepModel> moItemStepList = MoStepListTemp;

            MoStepGridView.AutoGenerateColumns = false;
            MoStepBindingSource.DataSource = moItemStepList;
            MoStepGridView.DataSource = MoStepBindingSource;

            MoStepGridView.Refresh();
        }

        public void ShowCheckedMoStepGridViewData()
        {

            List<MoItemStepModel> moItemStepList = MoStepListCheckTemp;

            MoStepGridView.AutoGenerateColumns = false;
            MoStepBindingSource.DataSource = moItemStepList;
            MoStepGridView.DataSource = MoStepBindingSource;

            MoStepGridView.Refresh();
        }

        private void Compute_Click(object sender, EventArgs e)
        {
            MoIdListLocal.Clear();

            Compute.Enabled = false;

            Thread thread = new Thread(() => {
                try
                {
                    List<string> MoIdList = new List<string>();
                    var dataSelect = MoGridView.SelectedRows;
                    foreach (DataGridViewRow row in dataSelect)
                    {
                        MoIdList.Add((string)row.Cells[0].Value);                       
                    }

                    MoIdListLocal = MoIdList;

                    MoStepListCheckTemp = SchedulePlan.DoSchedule(MoIdList);

                    //GetMoStepGridViewData();

                    Invoke(computerDelegate);
                }
                catch (Exception ex)
                {
                        
                    Invoke(errorDelegate,ex.Message);
                }
            });

            thread.IsBackground = true;
            thread.Start();

            

        }

        private void Refresh_Click(object sender, EventArgs e)
        {
            Refresh.Enabled = false;

            MoGridPageNumber = 0;
            MoItemGridPageNumber = 0;

            MoIdListLocal.Clear();

            Thread thread = new Thread(() => {
                BuildSchedulePlan();
                GetMoGridViewData();
                GetMoStepGridViewData();
                Invoke(refreshDelegate);
            });

            thread.IsBackground = true;
            thread.Start();
        }

        private void AdjustPreviousBtn_Click(object sender, EventArgs e)
        {

        }

        private void AdjustNextBtn_Click(object sender, EventArgs e)
        {

        }

        private void SaveBtn_Click(object sender, EventArgs e)
        {

            SchedulePlan.SaveChange(MoIdListLocal);

            Refresh_Click(sender, e);
        }

        private void BuildSchedulePlan()
        {
            SchedulePlan = new SchedulePlan(OrBitUserId);
        }

        private void MoGridView_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;           

            if (dgv.Columns[e.ColumnIndex].SortMode == DataGridViewColumnSortMode.Programmatic)
            {
                string columnBindingName = dgv.Columns[e.ColumnIndex].DataPropertyName;
                switch (dgv.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection)
                {
                    case System.Windows.Forms.SortOrder.None:
                        CustomSort(columnBindingName, "asc");
                        dgv.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection = System.Windows.Forms.SortOrder.Ascending;
                        break;
                    case System.Windows.Forms.SortOrder.Ascending:
                        CustomSort(columnBindingName, "desc");
                        dgv.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection = System.Windows.Forms.SortOrder.Descending;                       
                        break;
                    case System.Windows.Forms.SortOrder.Descending:
                        CustomSort(columnBindingName, "");
                        dgv.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection = System.Windows.Forms.SortOrder.None;
                        break;
                    
                }
            }

            foreach (DataGridViewColumn colum in dgv.Columns)
            {
                if(colum.Index != e.ColumnIndex)
                    colum.HeaderCell.SortGlyphDirection = System.Windows.Forms.SortOrder.None;
            }
        }

        private void CustomSort(string columnBindingName, string sortMode)
        {
            var bindsource = this.MoGridView.DataSource as BindingSource;

            var list = bindsource.DataSource as List<MoModel>;

            if(!string.IsNullOrEmpty(sortMode))
            {            
                list.Sort(
                delegate (MoModel info1, MoModel info2)
                {
                    Type t1 = info1.GetType();
                    Type t2 = info2.GetType();
                    PropertyInfo pro1 = t1.GetProperty(columnBindingName);
                    PropertyInfo pro2 = t2.GetProperty(columnBindingName);
                    return sortMode.ToLower().Equals("asc") ?
                    pro1.GetValue(info1, null).ToString().CompareTo(pro2.GetValue(info2, null).ToString()) :
                    pro2.GetValue(info2, null).ToString().CompareTo(pro1.GetValue(info1, null).ToString());
                });
            }
            else
            {
                list = list.OrderBy(p => p.SapDeliveryDate).ToList();
            }

            bindsource.DataSource = list;

            this.MoGridView.DataSource = bindsource;
            this.MoGridView.Refresh();
        }

        private void searchBtn_Click(object sender, EventArgs e)
        {
            string searchString = searchBox.Text;

            if (!string.IsNullOrEmpty(searchString))
            {
                var bindingsource = MoGridView.DataSource as BindingSource;

                var list = MoListTemp.Where(p => p.ProductName.Contains(searchString) || p.ProductId.Contains(searchString) || p.MOName.Contains(searchString) || p.MOId.Contains(searchString)).ToList();

                bindingsource.DataSource = list;
                MoGridView.DataSource = bindingsource;
            }
            else
            {
                //if(MoListTemp.Any())
                //{
                //    ShowMoGridViewData();                   
                //}
                //else
                //{
                MoBindingSource.DataSource = MoListTemp;
                MoGridView.DataSource = MoBindingSource;
                //}                
            }

            MoGridView.Refresh();
        }

        private void searchBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)//判断回车键
            {
                this.searchBtn_Click(sender, e);//触发按钮事件
            }
        }

        private void searchBtn2_Click(object sender, EventArgs e)
        {
            string searchString = searchBox2.Text;

            if (!string.IsNullOrEmpty(searchString))
            {
                var bindingsource = MoStepGridView.DataSource as BindingSource;

                var list = MoStepListTemp.Where(p => p.ProductName.Contains(searchString) || p.ProductId.Contains(searchString) || p.Specification.Contains(searchString) || p.MOName.Contains(searchString) || p.TaskStatus.Contains(searchString)).ToList();

                bindingsource.DataSource = list;
                MoStepGridView.DataSource = bindingsource;
            }
            else
            {
                //if(MoListTemp.Any())
                //{
                //    ShowMoGridViewData();                   
                //}
                //else
                //{
                MoStepBindingSource.DataSource = MoStepListTemp;
                MoStepGridView.DataSource = MoStepBindingSource;
                //}                
            }

            MoStepGridView.Refresh();
        }

        private void searchBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)//判断回车键
            {
                this.searchBtn2_Click(sender, e);//触发按钮事件
            }
        }

        private void MoStepGridView_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;

            if (dgv.Columns[e.ColumnIndex].SortMode == DataGridViewColumnSortMode.Programmatic)
            {
                string columnBindingName = dgv.Columns[e.ColumnIndex].DataPropertyName;
                switch (dgv.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection)
                {
                    case System.Windows.Forms.SortOrder.None:
                        CustomSort2(columnBindingName, "asc");
                        dgv.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection = System.Windows.Forms.SortOrder.Ascending;
                        break;
                    case System.Windows.Forms.SortOrder.Ascending:
                        CustomSort2(columnBindingName, "desc");
                        dgv.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection = System.Windows.Forms.SortOrder.Descending;
                        break;
                    case System.Windows.Forms.SortOrder.Descending:
                        CustomSort2(columnBindingName, "");
                        dgv.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection = System.Windows.Forms.SortOrder.None;
                        break;

                }
            }

            foreach (DataGridViewColumn colum in dgv.Columns)
            {
                if (colum.Index != e.ColumnIndex)
                    colum.HeaderCell.SortGlyphDirection = System.Windows.Forms.SortOrder.None;
            }
        }

        private void CustomSort2(string columnBindingName, string sortMode)
        {
            var bindsource = this.MoStepGridView.DataSource as BindingSource;

            var list = bindsource.DataSource as List<MoItemStepModel>;

            if (!string.IsNullOrEmpty(sortMode))
            {
                list.Sort(
                delegate (MoItemStepModel info1, MoItemStepModel info2)
                {
                    Type t1 = info1.GetType();
                    Type t2 = info2.GetType();
                    PropertyInfo pro1 = t1.GetProperty(columnBindingName);
                    PropertyInfo pro2 = t2.GetProperty(columnBindingName);
                    return sortMode.ToLower().Equals("asc") ?
                    pro1.GetValue(info1, null).ToString().CompareTo(pro2.GetValue(info2, null).ToString()) :
                    pro2.GetValue(info2, null).ToString().CompareTo(pro1.GetValue(info1, null).ToString());
                });
            }
            else
            {
                list = list.OrderBy(p => p.SapDeliveryDate).ToList();
            }

            bindsource.DataSource = list;

            this.MoStepGridView.DataSource = bindsource;
            this.MoStepGridView.Refresh();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox1.Checked)
            {
                //var bindingsource = MoStepGridView.DataSource as BindingSource;

                //MoStepListCheckTemp = bindingsource.DataSource as List<MoItemStepModel>;

                //var list = MoStepListCheckTemp.Where(p => p.TaskStatus.Contains("已登记")).ToList();

                //bindingsource.DataSource = list;
                //MoStepGridView.DataSource = bindingsource;


                ShowCheckedMoStepGridViewData();
            }
            else
            {
                //MoStepBindingSource.DataSource = MoStepListCheckTemp;
                //MoStepGridView.DataSource = MoStepBindingSource;

                ShowMoStepGridViewData();
            }

            MoStepGridView.Refresh();
        }

        public void computerCallBack()
        {
            
            checkBox1.CheckState = CheckState.Checked;
            MessageBox.Show("计算结束");

            Compute.Enabled = true;
        }

        public void outSideInvokeCallBack()
        {
            SaveBtn_Click(null,null);
        }

        public void refreshCallBack()
        {
            ShowMoGridViewData();

            checkBox1.Checked = false;
            checkBox1.CheckState = CheckState.Unchecked;

            ShowMoStepGridViewData();
            MessageBox.Show("刷新结束");

            Refresh.Enabled = true;
        }

        public void errorCallBack(string message)
        {
            Compute.Enabled = true;
            Refresh.Enabled = true;
            MessageBox.Show(message);
        }

        #region 甘特图
        OverlayPainter _mOverlay = new OverlayPainter();

        ProjectManager _mManager = null;

        public void BuildManager(Venture venture)
        {
            // Create a Project and some Tasks
            _mManager = new ProjectManager();

            var ventureTask = new MyTask(_mManager) { Name = venture.VentureName };
            _mManager.Add(ventureTask);

            List<MoItemStepModel> taskList = SchedulePlan.GetMoItemStopModelByVentureId(venture.VentureId, _mManager);

            foreach (var task in taskList)
            {
                _mManager.Group(ventureTask, task);
            }

            

            //var jake = new MyResource() { Name = "Jake", Name2 = "123456" };
            //var peter = new MyResource() { Name = "Peter" };
            //var john = new MyResource() { Name = "John" };
            //var lucas = new MyResource() { Name = "Lucas" };
            //var james = new MyResource() { Name = "James" };
            //var mary = new MyResource() { Name = "Mary" };
            // Add some resources
            //_mManager.Assign(wake, jake);
            //_mManager.Assign(wake, peter);
            //_mManager.Assign(wake, john);
            //_mManager.Assign(teeth, jake);
            //_mManager.Assign(teeth, james);
            //_mManager.Assign(pack, james);
            //_mManager.Assign(pack, lucas);
            //_mManager.Assign(shower, mary);
            //_mManager.Assign(shower, lucas);
            //_mManager.Assign(shower, john);

            // Initialize the Chart with our ProjectManager and CreateTaskDelegate


            //List<string> list = _mManager.ResourcesOf(wake).Select(x => (x as MyResource).Name).ToList();
            //list.AddRange(_mManager.ResourcesOf(wake).Select(x => (x as MyResource).Name2));
            //// set some tooltips to show the resources in each task
            //_mChart.SetToolTip(wake, string.Join(", ", list));
            //_mChart.SetToolTip(teeth, string.Join(", ", _mManager.ResourcesOf(teeth).Select(x => (x as MyResource).Name)));
            //_mChart.SetToolTip(pack, string.Join(", ", _mManager.ResourcesOf(pack).Select(x => (x as MyResource).Name)));
            //_mChart.SetToolTip(shower, string.Join(", ", _mManager.ResourcesOf(shower).Select(x => (x as MyResource).Name)));

            // Set Time information

        }

        public void InitGanttChart()
        {
            _mChart.Init(_mManager);
            _mChart.CreateTaskDelegate = delegate () { return new MyTask(_mManager); };

            // Attach event listeners for events we are interested in
            _mChart.TaskMouseOver += new EventHandler<TaskMouseEventArgs>(_mChart_TaskMouseOver);
            _mChart.TaskMouseOut += new EventHandler<TaskMouseEventArgs>(_mChart_TaskMouseOut);
            _mChart.TaskSelected += new EventHandler<TaskMouseEventArgs>(_mChart_TaskSelected);
            //_mChart.PaintOverlay += _mOverlay.ChartOverlayPainter;
            _mChart.AllowTaskDragDrop = false;

            var span = DateTime.Today - _mManager.Start;
            _mManager.Now = span; // set the "Now" marker at the correct date
            _mChart.TimeResolution = TimeResolution.Day; // Set the chart to display in days in header
        }

        public void GanttCallBack()
        {
            _mChart.Init(_mManager);
            _mChart.Invalidate();
        }

        void _mChart_TaskSelected(object sender, TaskMouseEventArgs e)
        {
            //_mTaskGrid.SelectedObjects = _mChart.SelectedTasks.Select(x => _mManager.IsPart(x) ? _mManager.SplitTaskOf(x) : x).ToArray();
            //_mResourceGrid.Items.Clear();
            //_mResourceGrid.Items.AddRange(_mManager.ResourcesOf(e.Task).Select(x => new ListViewItem(((MyResource)x).Name)).ToArray());
        }

        void _mChart_TaskMouseOut(object sender, TaskMouseEventArgs e)
        {
            //lblStatus.Text = "";
            //_mChart.Invalidate();
        }

        void _mChart_TaskMouseOver(object sender, TaskMouseEventArgs e)
        {
            //lblStatus.Text = string.Format("{0} to {1}", _mManager.GetDateTime(e.Task.Start).ToLongDateString(), _mManager.GetDateTime(e.Task.End).ToLongDateString());
            //_mChart.Invalidate();
        }

        #endregion

        #region overlay painter
        /// <summary>
        /// An example of how to encapsulate a helper painter for painter additional features on Chart
        /// </summary>
        public class OverlayPainter
        {
            /// <summary>
            /// Hook such a method to the chart paint event listeners
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            public void ChartOverlayPainter(object sender, Braincase.GanttChart.ChartPaintEventArgs e)
            {
                // Don't want to print instructions to file
                if (this.PrintMode) return;

                var g = e.Graphics;
                var chart = e.Chart;

                // Demo: Static billboards begin -----------------------------------
                // Demonstrate how to draw static billboards
                // "push matrix" -- save our transformation matrix
                e.Chart.BeginBillboardMode(e.Graphics);

                // draw mouse command instructions
                int margin = 300;
                int left = 20;
                var color = chart.HeaderFormat.Color;
                StringBuilder builder = new StringBuilder();
                builder.AppendLine("THIS IS DRAWN BY A CUSTOM OVERLAY PAINTER TO SHOW DEFAULT MOUSE COMMANDS.");
                builder.AppendLine("*******************************************************************************************************");
                builder.AppendLine("Left Click - Select task and display properties in PropertyGrid");
                builder.AppendLine("Left Mouse Drag - Change task starting point");
                builder.AppendLine("Right Mouse Drag - Change task duration");
                builder.AppendLine("Middle Mouse Drag - Change task complete percentage");
                builder.AppendLine("Left Doubleclick - Toggle collaspe on task group");
                builder.AppendLine("Right Doubleclick - Split task into task parts");
                builder.AppendLine("Left Mouse Dragdrop onto another task - Group drag task under drop task");
                builder.AppendLine("Right Mouse Dragdrop onto another task part - Join task parts");
                builder.AppendLine("SHIFT + Left Mouse Dragdrop onto another task - Make drop task precedent of drag task");
                builder.AppendLine("ALT + Left Dragdrop onto another task - Ungroup drag task from drop task / Remove drop task from drag task precedent list");
                builder.AppendLine("SHIFT + Left Mouse Dragdrop - Order tasks");
                builder.AppendLine("SHIFT + Middle Click - Create new task");
                builder.AppendLine("ALT + Middle Click - Delete task");
                builder.AppendLine("Left Doubleclick - Toggle collaspe on task group");
                var size = g.MeasureString(builder.ToString(), e.Chart.Font);
                var background = new Rectangle(left, chart.Height - margin, (int)size.Width, (int)size.Height);
                background.Inflate(10, 10);
                g.FillRectangle(new System.Drawing.Drawing2D.LinearGradientBrush(background, Color.LightYellow, Color.Transparent, System.Drawing.Drawing2D.LinearGradientMode.Vertical), background);
                g.DrawRectangle(Pens.Brown, background);
                g.DrawString(builder.ToString(), chart.Font, color, new PointF(left, chart.Height - margin));


                // "pop matrix" -- restore the previous matrix
                e.Chart.EndBillboardMode(e.Graphics);
                // Demo: Static billboards end -----------------------------------
            }

            public bool PrintMode { get; set; }
        }
        #endregion overlay painter

        #region custom task and resource
        /// <summary>
        /// A custom resource of your own type (optional)
        /// </summary>
        [Serializable]
        public class MyResource
        {
            public string Name { get; set; }
            public string Name2 { get; set; }
        }
        /// <summary>
        /// A custom task of your own type deriving from the Task interface (optional)
        /// </summary>
        [Serializable]
        public class MyTask : Task
        {
            public MyTask(ProjectManager manager)
                : base()
            {
                Manager = manager;
            }

            private ProjectManager Manager { get; set; }

            public new TimeSpan Start { get { return base.Start; } set { Manager.SetStart(this, value); } }
            public new TimeSpan End { get { return base.End; } set { Manager.SetEnd(this, value); } }
            public new TimeSpan Duration { get { return base.Duration; } set { Manager.SetDuration(this, value); } }
            public new float Complete { get { return base.Complete; } set { Manager.SetComplete(this, value); } }
        }
        #endregion custom task and resource

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Thread thread = new Thread(() => {
            //    try
            //    {
            //        if(currentVenture != null)
            //        {
            //            InitGanttChart(currentVenture);

            //            Invoke(ganttDelegate);
            //        }
            //    }
            //    catch (Exception ex)
            //    {

            //        Invoke(errorDelegate, ex.Message);
            //    }
            //});

            //thread.IsBackground = true;
            //thread.Start();

            ListBox listbox = (ListBox)sender;

            Venture venture = (Venture)listbox.SelectedItem;

            currentVenture = venture;


            if (currentVenture != null)
            {
                BuildManager(currentVenture);

                GanttCallBack();
            }
        }

        private void MoStepGridView_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
            {

                if (e.NewValue + MoStepGridView.DisplayedRowCount(false) == MoStepGridView.Rows.Count)
                {

                    //MessageBox.Show(string.Format("NewValue:{0}--OldValue:{1}--DisplayedRowCount:{2}", e.NewValue, e.OldValue, MoStepGridView.DisplayedRowCount(false)));
                    //MessageBox.Show("到底了，可以加载新数据了！");
                    //这里面写加载数据的相关操作逻辑

                    MoItemGridPageNumber++;

                    //MoStepListTemp.AddRange(SchedulePlan.GetScheduledMoItemStepModelList(MoItemGridPageNumber));

                    if(!checkBox1.Checked)
                    {
                        var list = SchedulePlan.GetScheduledMoItemStepModelList(MoItemGridPageNumber);

                        foreach (var l in list)
                        {
                            MoStepBindingSource.Add(l);
                        }
                    }                   
                }
            }
        }

        private void MoGridView_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
            {

                if (e.NewValue + MoGridView.DisplayedRowCount(false) == MoGridView.Rows.Count)
                {

                    //MessageBox.Show(string.Format("NewValue:{0}--OldValue:{1}--DisplayedRowCount:{2}", e.NewValue, e.OldValue, MoGridView.DisplayedRowCount(false)));
                    //MessageBox.Show("到底了，可以加载新数据了！");
                    //这里面写加载数据的相关操作逻辑

                    MoGridPageNumber++;

                    //MoListTemp.AddRange(SchedulePlan.GetMoList(MoGridPageNumber));

                    var list = SchedulePlan.GetMoList(MoGridPageNumber);

                    foreach (var l in list)
                    {
                        MoBindingSource.Add(l);
                    }

                }
            }
        }

        #region excel导入

        public List<List<object>> ReadExcel(Stream fileStream)
        {
            List<List<object>> beanList = new List<List<object>>();

            using (ExcelPackage package = new ExcelPackage(fileStream))
            {
                for (int i = 1; i <= package.Workbook.Worksheets.Count; ++i)
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets[i];
                    for (int m = sheet.Dimension.Start.Row, n = sheet.Dimension.End.Row; m <= n; m++)
                    {
                        List<object> objectList = new List<object>();
                        for (int j = sheet.Dimension.Start.Column, k = sheet.Dimension.End.Column; j <= k; j++)
                        {
                            object o = sheet.Cells[m, j].Value;
                            objectList.Add(o);
                        }
                        beanList.Add(objectList);
                    }
                }
            }
            beanList.RemoveAt(0);

            return beanList;
        }

        public List<MOItemTask> ConvertToMOItemTask(List<List<object>> beanList)
        {
            List<MOItemTask> pList = new List<MOItemTask>();

            foreach (var bean in beanList)
            {
                try
                {

                    MOItemTask p = new MOItemTask();

                    if (bean[0] != null)
                        p.MOItemTaskName = (string)bean[0];
                    else
                        continue;                                       

                    try
                    {
                        if (bean[1] != null)
                        {
                            //string strDate = DateTime.FromOADate(Convert.ToDouble(bean[1])).ToString("d");

                            p.PlannedDateFrom = DateTime.FromOADate(Convert.ToDouble(bean[1]));//Convert.ToDateTime(strDate);
                        }
                        //p.DELIVERY_DATE = Convert.ToDateTime(bean[11]);
                        else
                            continue;
                    }
                    catch
                    {
                        if (bean[1] != null)
                            p.PlannedDateFrom = (DateTime)bean[1];
                        else
                            continue;
                    }

                    try
                    {
                        if (bean[2] != null)
                        {
                            //string strDate = DateTime.FromOADate(Convert.ToDouble(bean[2])).ToString("d");

                            p.PlannedDateTo = DateTime.FromOADate(Convert.ToDouble(bean[2]));//Convert.ToDateTime(strDate);
                        }
                        //p.DELIVERY_DATE = Convert.ToDateTime(bean[11]);
                        else
                            continue;
                    }
                    catch
                    {
                        if (bean[2] != null)
                            p.PlannedDateTo = (DateTime)bean[2];
                        else
                            continue;
                    }


                    if (bean[3] != null)
                        p.VentureId = (string)bean[3];
                    else
                        continue;


                    pList.Add(p);
                }
                catch (Exception e)
                {
                    e.ToString();
                    continue;
                }
            }

            return pList;
        }

        #endregion

        private void FileUploadBtn_Click(object sender, EventArgs e)
        {
            //初始化一个OpenFileDialog类
            OpenFileDialog fileDialog = new OpenFileDialog();

            //判断用户是否正确的选择了文件
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                //获取用户选择文件的后缀名
                string extension = Path.GetExtension(fileDialog.FileName);
                //声明允许的后缀名
                string[] str = new string[] { ".xls", ".xlsx" };
                if (!str.Contains(extension))
                {
                    MessageBox.Show("仅支持导入Excel文件！");
                }
                else
                {
                    //获取用户选择的文件
                    FileInfo fileInfo = new FileInfo(fileDialog.FileName);

                    List<List<object>> objectList = ReadExcel(fileInfo.OpenRead());

                    List<MOItemTask> mitList = ConvertToMOItemTask(objectList);

                    var retstring = SchedulePlan.ExcelImport(mitList);

                    MessageBox.Show(retstring);

                    Refresh_Click(sender, e);
                }
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            DateTime start = dateTimePicker1.Value;
            DateTime end = dateTimePicker2.Value;


            List<ReportModel> reportModelList = SchedulePlan.GetSpecificationCapReport(MoIdListLocal, start,end);

            ChartArea mainArea = new ChartArea();
            ChartArea mainArea2 = new ChartArea();
            mainArea.Axes.FirstOrDefault().IntervalAutoMode = IntervalAutoMode.VariableCount;
            mainArea.Axes.FirstOrDefault().LabelStyle.IsStaggered = true;
            mainArea.AxisX.Interval = 1;   //设置X轴坐标的间隔为1
            mainArea.AxisX.IntervalOffset = 1;  //设置X轴坐标偏移为1
            chart1.ChartAreas.Clear();
            chart1.ChartAreas.Add(mainArea);

            Series s1 = new Series();
            Series s2 = new Series();
            Series s3 = new Series();



            foreach (var reportModel in reportModelList)
            {
                s1.Points.AddXY(reportModel.Name, reportModel.Plan);
                s2.Points.AddXY(reportModel.Name, reportModel.Real);
                s3.Points.AddXY(reportModel.Name, reportModel.Rate);
            }

            s1.Color = Color.Green;
            s2.Color = Color.Red;
            s3.Color = Color.Blue;

            s1.LegendText = "能力工时";
            s2.LegendText = "下达工时";
            s3.LegendText = "排产工时";
            s1.IsValueShownAsLabel = true;
            s2.IsValueShownAsLabel = true;
            s3.IsValueShownAsLabel = true;

            chart1.Series.Clear();

            chart1.Series.Add(s1);
            chart1.Series.Add(s2);
            chart1.Series.Add(s3);

            
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            SaveBtn_Click(sender, e);
        }

        private void TabPage3_Click(object sender, EventArgs e)
        {  
            
            
        }

        private void Chart1_Click(object sender, EventArgs e)
        {
            
            HitTestResult Result = new HitTestResult();
            Result = chart1.HitTest(((MouseEventArgs)e).X, ((MouseEventArgs)e).Y);
            if (Result.Series != null)
            {
                //MessageBox.Show("'X轴:'" + Result.Series.Points[Result.PointIndex].XValue.ToString() + "'Y轴:'" + Result.Series.Points[Result.PointIndex].YValues[0].ToString());
                string specificationName = Result.Series.Points[Result.PointIndex].AxisLabel;

                DateTime start = dateTimePicker1.Value;
                DateTime end = dateTimePicker2.Value;


                List<ReportModel> reportModelList = SchedulePlan.GetSpecificationCapReportDay(MoIdListLocal, start, end, specificationName);

                ChartArea mainArea = new ChartArea();
                mainArea.Axes.FirstOrDefault().IntervalAutoMode = IntervalAutoMode.VariableCount;
                mainArea.Axes.FirstOrDefault().LabelStyle.IsStaggered = true;
                mainArea.AxisX.Interval = 1;   //设置X轴坐标的间隔为1
                mainArea.AxisX.IntervalOffset = 1;  //设置X轴坐标偏移为1
                chart2.ChartAreas.Clear();
                chart2.ChartAreas.Add(mainArea);

                Series s1 = new Series();
                Series s2 = new Series();
                Series s3 = new Series();



                foreach (var reportModel in reportModelList)
                {
                    s1.Points.AddXY(reportModel.Name, reportModel.Plan);
                    s2.Points.AddXY(reportModel.Name, reportModel.Real);
                    s3.Points.AddXY(reportModel.Name, reportModel.Rate);
                }

                s1.Color = Color.Green;
                s2.Color = Color.Red;
                s3.Color = Color.Blue;

                s1.LegendText = "能力工时";
                s2.LegendText = "下达工时";
                s3.LegendText = "排产工时";
                s1.IsValueShownAsLabel = true;
                s2.IsValueShownAsLabel = true;
                s3.IsValueShownAsLabel = true;

                chart2.Series.Clear();

                chart2.Series.Add(s1);
                chart2.Series.Add(s2);
                chart2.Series.Add(s3);

                chart2.Titles.Add(specificationName);

            }

                
        }
    }


}
