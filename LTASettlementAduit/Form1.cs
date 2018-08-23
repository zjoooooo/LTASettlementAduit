using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using log4net;

[assembly: log4net.Config.DOMConfigurator(ConfigFileExtension = "config", Watch = true)]

namespace LTASettlementAduit
{
    public partial class Form1 : Form
    {
        public ArrayList al = new ArrayList();
        private string alarmTxt = "";
        Dictionary<string, string> carparklist = new Dictionary<string, string>();
        Dictionary<string, string> batchlist = new Dictionary<string, string>();
        Dictionary<string, string> batchIP = new Dictionary<string, string>();
        static ILog log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        public Form1()
        {
            InitializeComponent();

        }
        private void InitCarparkList()
        {
            string constr = "Data Source=172.16.1.89;uid=secure;pwd=weishenme;database=carpark";
            string CommandText = @"select name,ip,batch from Whole where valid =1";
            //           string CommandText = @"select name,ip,batch from Whole where name='HG10'";
            DataSet ds = null;

            try
            {
                ds = SqlHelper.ExecuteDataset(constr, CommandType.Text, CommandText);
                foreach (DataRow ls in ds.Tables[0].Rows)
                {
                    carparklist.Add(ls[0].ToString(), ls[1].ToString());
                    batchlist.Add(ls[0].ToString(), ls[2].ToString());
                }
            }
            catch (SqlException sql)
            {
                log.Error($"Fail To Get Car Park List!{sql.ToString()}");
            }
            finally
            {
                try
                {
                    if (ds != null)
                        ds.Dispose();
                }
                catch (SqlException e)
                {
                    log.Error("Fail To Close Car Park List DataSet:" + e.ToString());
                }
            }

            string cmd_ = "SELECT* FROM [dbo].[ServerDetails]";
            string constr_ = "Data Source=172.16.1.89;uid=secure;pwd=weishenme;database=LTASettlementAudit";

            DataSet ds_ = null;

            try
            {
                ds_ = SqlHelper.ExecuteDataset(constr_, CommandType.Text, cmd_);
                foreach (DataRow ls in ds_.Tables[0].Rows)
                {
                    batchIP.Add(ls[0].ToString(), ls[1].ToString());
                }
            }
            catch (SqlException sql)
            {
                log.Error($"Fail To Get Car Park List!{sql.ToString()}");
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            InitCarparkList();
            if (!GetValue("manual").Equals("true"))
            {
                Thread thr = new Thread(() => RunCepassCheck());
                thr.Start();
            }

        }
        #region Cepass file upload check.
        private void AlarmTxt(string str)
        {
            alarmTxt = alarmTxt + Environment.NewLine;
            alarmTxt = alarmTxt + str;
            alarmTxt = alarmTxt + Environment.NewLine;
        }
        private static string GetValue(string strkey)
        {

            foreach (string key in ConfigurationManager.AppSettings)
            {
                if (key.Contains(strkey))
                {
                    return ConfigurationManager.AppSettings[key];
                }
            }
            return "";

        }

        private static List<string> GetEmail(string strkey)
        {
            List<string> list = new List<string>();
            foreach (string key in ConfigurationManager.AppSettings)
            {
                if (key.Contains(strkey))
                {
                    list.Add(ConfigurationManager.AppSettings[key]);
                }
            }
            return list;
        }
        private void SendEmail(string sub, string body, string address)
        {
            // Command line argument must the the SMTP host.
            SmtpClient client = new SmtpClient();
            client.Port = 587;
            client.Host = "smtp.gmail.com";
            client.EnableSsl = true;
            client.Timeout = 15000;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.Credentials = new NetworkCredential("seasonalarm@gmail.com", "wei3shen2me");
            MailMessage mm = new MailMessage("seasonalarm@gmail.com", address, sub, body);
            MailAddress copy1 = new MailAddress("jzhang@Secureparking.com.sg");
           // MailAddress copy2 = new MailAddress("leon@Secureparking.com.sg");
           // MailAddress copy3 = new MailAddress("schew@secureparking.com.sg");
            mm.CC.Add(copy1);    //CC email 
           // mm.CC.Add(copy2);
           // mm.CC.Add(copy3);
            mm.BodyEncoding = UTF8Encoding.UTF8;
            mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;

            List<string> lis = GetEmail("email");

            if (lis != null)
            {
                foreach (string str in lis)
                {
                    mm.To.Add(new MailAddress(str, "seasonalarm@gmail.com"));
                }
            }

            try
            {
                client.Send(mm);
                LogClass.WriteLog("Mail Sent! Success");
            }
            catch (Exception ex)
            {
                LogClass.WriteLog(ex.ToString());
            }
        }
        private void CleanDB()
        {
            string constr_server = "Data Source=172.16.1.89;uid=secure;pwd=weishenme;database=LTASettlementAudit";
            string cmd = @"Delete [dbo].[Collection_File_History] where Coll_Create_dt BETWEEN @start_time and @end_time";
            string start_time = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
            string end_time = DateTime.Now.ToString("yyyy-MM-dd 00:00:00");

            SqlParameter[] para = new SqlParameter[]
             {
                    new SqlParameter("@start_time",start_time),
                    new SqlParameter("@end_time",end_time)
             };

            try
            {
                SqlHelper.ExecuteNonQuery(constr_server, CommandType.Text, cmd,para);
            }
            catch (SqlException e)
            {
                LogClass.WriteLog($"read server Collection_File_History error {e.ToString()}");
                AlarmTxt($"read server Collection_File_History error");
                return;
            }
        }
        private void ReadDataFromPMS()
        {
            string constr_server = "Data Source=172.16.1.89;uid=secure;pwd=weishenme;database=LTASettlementAudit";
            string start_time = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
            string end_time = DateTime.Now.ToString("yyyy-MM-dd 00:00:00");
            foreach (KeyValuePair<string, string> kv in carparklist)
            {
                SqlBulkCopy sbc = null;
                DataColumn CarparkColumn = null;
                DataColumn BatchColumn = null;
                string constr = $"Data Source={kv.Value};uid=sa;pwd=yzhh2007;database={kv.Key}";
                string cmd = @"SELECT Collection_File,Coll_Create_dt,Total_Trans,Total_Amt,Send_Flag,Send_dt,Control_File,Control_Create_dt,Add_dt,Update_dt FROM [dbo].[Collection_File_History] where Coll_Create_dt BETWEEN @start_time and @end_time;";
                DataSet ds = null;
                SqlParameter[] para = new SqlParameter[]
                {
                    new SqlParameter("@start_time",start_time),
                    new SqlParameter("@end_time",end_time)
                };

                try
                {
                    ds = SqlHelper.ExecuteDataset(constr, CommandType.Text, cmd, para);
                    LogClass.WriteLog($"{kv.Key} Collected LTA Data.");
                }
                catch (SqlException e)
                {
                    LogClass.WriteLog($"{kv.Key} read db fail,{e.ToString()}");
                    AlarmTxt($"{kv.Key} Fail To Collect Cepass File Details,Please Check PMS Connection.");
                    continue;
                }

                if (ds == null)
                {
                    LogClass.WriteLog($"{kv.Key} dataset is null");
                    AlarmTxt($"{kv.Key} Dataset Is Null");
                    continue;
                }

                try
                {
                    if ((ds != null) && (ds.Tables[0].Rows.Count > 0))
                    {
                        CarparkColumn = new DataColumn();
                        CarparkColumn.DataType = System.Type.GetType("System.String");
                        CarparkColumn.ColumnName = "carparkID";
                        CarparkColumn.DefaultValue = kv.Key;

                        BatchColumn = new DataColumn();
                        BatchColumn.DataType = System.Type.GetType("System.String");
                        BatchColumn.ColumnName = "batch";
                        BatchColumn.DefaultValue = batchlist[kv.Key];

                        ds.Tables[0].Columns.Add(CarparkColumn);
                        ds.Tables[0].Columns.Add(BatchColumn);
                        sbc = new SqlBulkCopy(constr_server);
                        sbc.DestinationTableName = "Collection_File_History";
                        sbc.WriteToServer(ds.Tables[0]);
                    }
                }

                catch (Exception e)
                {
                    LogClass.WriteLog($"copy error {e.ToString()}");
                    AlarmTxt($"{kv.Key} Column Copy Error");
                }
            }
        }
        private void compareWithServer()
        {
            string constr_server = "Data Source=172.16.1.89;uid=secure;pwd=weishenme;database=LTASettlementAudit";
            string cmd = @"SELECT * FROM [dbo].[Collection_File_History] where Coll_Create_dt BETWEEN @start_time and @end_time ORDER BY batch";
            string start_time = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
            string end_time = DateTime.Now.ToString("yyyy-MM-dd 00:00:00");
            SqlParameter[] para = new SqlParameter[]
             {
                    new SqlParameter("@start_time",start_time),
                    new SqlParameter("@end_time",end_time)
             };
            SqlDataReader reader = null;
            try
            {
                    reader = SqlHelper.ExecuteReader(constr_server, CommandType.Text, cmd, para);
            }
            catch (SqlException e)
            {
                LogClass.WriteLog($"Reading Error Collection_File_History From Server {e.ToString()}");
                AlarmTxt($"Reading Error Collection_File_History From Server");
                return;
            }
            while ((reader != null) && (reader.Read()))
            {
                string collec_file = reader["Collection_File"].ToString();
                string carpark = reader["carparkID"].ToString();
                string batch = reader["batch"].ToString();
                string constr_batch_server = null;

                try
                {
                    constr_batch_server = $"Data Source={batchIP[batch]};uid=sa;pwd=yzhh2007;database=LTACS";
                }
                catch (Exception e)
                {
                    AlarmTxt($"Program Can Not Find LTA Server For Batch {batch},{e.ToString()}");
                }

                string cmd_batch_server = $"SELECT * FROM [dbo].[Coll_File_History] where Upload_Time>'{DateTime.Now.ToString("yyyy-MM-dd")}' AND File_Name like '%{collec_file}'";
                DataSet ds = null;
                try
                {

                    ds = SqlHelper.ExecuteDataset(constr_batch_server, CommandType.Text, cmd_batch_server);
                   
                }catch(SqlException sqlec)
                {
                    LogClass.WriteLog($"Compare Collection File {collec_file} Fail At Server For {carpark},{sqlec.ToString()}");
                    AlarmTxt($"Compare Collection File {collec_file} Fail At Server For {carpark}");
                    continue;
                }

                if ((ds!=null)&&(ds.Tables[0].Rows.Count > 0))
                {
                    LogClass.WriteLog($"Collection File {collec_file} Upload Successfully For {carpark}");
                }
                else
                {
                    LogClass.WriteLog($"Collection File {collec_file} Never Upload For {carpark}");
                    AlarmTxt($"Collection File {collec_file} Never Upload For {carpark}");
                }
            }
            if((alarmTxt == null)||(alarmTxt==""))
            {
                alarmTxt = "All LTA Settlement Files already upload!!";
            }
            SendEmail("Daily check for LTA Settlement", alarmTxt, "leon@secureparking.com.sg");
        }
        private void RunCepassCheck()
        {
            CleanDB();
            ReadDataFromPMS();
            compareWithServer();
            Application.Exit();         
        }
        #endregion
        private void ReadPMS_Click(object sender, EventArgs e)
        {

            Thread thr = new Thread(() => ReadPMS());
            thr.Start();
        }
        private void ReadPMS()
        {
            CleanDB();
            ReadDataFromPMS();
        }
        private void CompareServer_Click(object sender, EventArgs e)
        {
            Thread thr = new Thread(() => compareWithServer());
            thr.Start();
        }

        #region Read Cepass Result File 
        private void CollectData()
        {
            string date = DateTime.Now.AddDays(-1).ToString();
            string CommandText = @"select name,ip,batch from Whole";
            foreach (KeyValuePair<string, string> kv in carparklist)
            {
                string connectString = "Data Source=" + kv.Value + ";uid=sa;pwd=yzhh2007;database=" + kv.Key;


            }
        }
        //Read log insert into db.
        private void ReadStreamLog(string path)
        {
            StreamReader srReadFile = new StreamReader(path);
            LogClass.WriteLog("reading.");
            // 读取流直至文件末尾结束
            while (!srReadFile.EndOfStream)
            {
                string strReadLine = srReadFile.ReadLine(); //读取每行数据

                if ((strReadLine == null) || (strReadLine == "null"))
                {
                    continue;
                }

                string constr = "Data Source=172.16.1.89;uid=secure;pwd=weishenme;database=LTASettlementAudit";
                string cmd = @"insert into RejectDetails(RejectDate,CashCardNo,Amount,Reason,CarPark)VALUES(@RejectDate,@CashCardNo,@Amount,@Reason,@CarPark)";
                string[] array = strReadLine.Split(',');
                string time = array[0];
                string card = array[1];
                string amount = array[2];
                string reason = array[3];
                string carpark = array[4];
                // LogClass.WriteLog($"time = {time}, card={card},amount={amount},reason={reason},carpark={carpark}");
                //Convert.ToDateTime(time).ToString("yyyy-MM-dd HH:mm:ss")

                string day = time.Substring(0, 2);  //day
                string month = time.Substring(3, 2);  // month
                string year = time.Substring(6, 4);  // year
                string hour = time.Substring(11, 8);  //hours

                time = $"{year}-{month}-{day} {hour}";

                //LogClass.WriteLog($"{year},{month},{day},{hour}");

                SqlParameter[] para = new SqlParameter[]
                {
                    new SqlParameter("@RejectDate",time),
                    new SqlParameter("@CashCardNo",card),
                    new SqlParameter("@Amount",amount),
                    new SqlParameter("@Reason",reason),
                    new SqlParameter("@CarPark",carpark)
                };

                try
                {
                    SqlHelper.ExecuteNonQuery(constr, CommandType.Text, cmd, para);
                    LogClass.WriteLog($"Insert ok for {card}");
                }
                catch (SqlException sql)
                {
                    LogClass.WriteLog($"Fail To Get Car Park List!{sql.ToString()}");
                }
            }

            Application.Exit();
        }
        //Read result file and put into log.
        private void ReadFromResultFiles()
        {
            GetAllDirList($"C:\\Users\\admin\\Desktop\\LTA");
            for (int i = 0; i < al.Count; i++)
            {
                DirectoryInfo folder = new DirectoryInfo(al[i].ToString());

                foreach (FileInfo file in folder.GetFiles("*.txt"))
                {
                    //HG
                    string[] HGstr = file.FullName.Split('\\');
                    string carpark = HGstr[6];
                    ReadStream(file.FullName, carpark);
                }
            }
            Application.Exit();
        }
        public void GetAllDirList(string strBaseDir)
        {

            DirectoryInfo di = new DirectoryInfo(strBaseDir);
            DirectoryInfo[] diA = di.GetDirectories();
            for (int i = 0; i < diA.Length; i++)

            {
                al.Add(diA[i].FullName);
                //diA[i].FullName是某个子目录的绝对地址，把它记录在ArrayList中
                GetAllDirList(diA[i].FullName);
                //注意：递归了。逻辑思维正常的人应该能反应过来
            }
        }
        private bool ReadStream(string Filename, string carpark)
        {
            // 读取文件的源路径及其读取流
            StreamReader srReadFile = new StreamReader(Filename);

            // 读取流直至文件末尾结束
            while (!srReadFile.EndOfStream)
            {
                string strReadLine = srReadFile.ReadLine(); //读取每行数据
                if ((strReadLine == null) || (strReadLine == ""))
                {
                    continue;
                }
                string[] ContentToken = strReadLine.Split(',');//截取数据
                string final = null;
                int count = 0;
                foreach (string str in ContentToken)
                {
                    string[] strArray = null;
                    if (count == 0)
                    {
                        strArray = str.Split(new string[] { "time:" }, StringSplitOptions.RemoveEmptyEntries);
                    }
                    else
                    {

                        strArray = str.Split(':');
                    }
                    final += strArray[1] + ",";
                    count++;
                }
                LogClass.WriteLog(final + carpark);


            }
            // 关闭读取流文件
            srReadFile.Close();
            return false;
        }
        //export EPLB EPLA VIP GRO
        private void Export()
        {
            string cmd = @"SELECT * FROM [dbo].[season_mst] where holder_type in (1,4,12,13);";
            foreach (KeyValuePair<string, string> kv in carparklist)
            {
                string constr = $"Data Source={kv.Value};uid=sa;pwd=yzhh2007;database={kv.Key}";

                DataSet ds = null;
                try
                {
                    ds = SqlHelper.ExecuteDataset(constr, CommandType.Text, cmd);
                }
                catch (SqlException sqle)
                {
                    LogClass.WriteLog($"fail to excute data, {sqle.ToString()}");
                    continue;
                }
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Application.Workbooks.Add(true);
                excel.Cells[1, 1] = "season_no";
                excel.Cells[1, 2] = "vehicle_no";
                excel.Cells[1, 3] = "date_from";
                excel.Cells[1, 4] = "date_to";
                excel.Cells[1, 5] = "holder_type";
                excel.Cells[1, 6] = "s_status";

                Microsoft.Office.Interop.Excel.Range Header = excel.get_Range("A1", "F1");// set bold
                Header.Font.Bold = true;
                Microsoft.Office.Interop.Excel.Range Whole_Sheet = excel.get_Range("A1", "F10000");
                Whole_Sheet.NumberFormatLocal = "@";
                int i = 1;
                if ((ds == null) || (ds.Tables[0].Rows.Count <= 0))
                {
                    continue;
                }
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    string IU = dr["season_no"].ToString();
                    string Vehicle_no = dr["vehicle_no"].ToString();
                    string Valid_from = dr["date_from"].ToString();
                    string Valid_To = dr["date_to"].ToString();
                    string Holder_type = dr["holder_type"].ToString();
                    string Valid_status = dr["s_status"].ToString();
                    switch (Holder_type)
                    {
                        case "1":
                            Holder_type = "VIP";
                            break;
                        case "4":
                            Holder_type = "GRO";
                            break;
                        case "12":
                            Holder_type = "EPLA";
                            break;
                        case "13":
                            Holder_type = "EPLB";
                            break;
                    }
                    switch (Valid_status)
                    {
                        case "1":
                            Valid_status = "Valid";
                            break;
                        case "2":
                            Valid_status = "Expired";
                            break;
                        case "3":
                            Valid_status = "Terminated";
                            break;
                        case "0":
                            Valid_status = "Invalid";
                            break;
                    }
                    excel.Cells[i + 1, 1] = IU;
                    excel.Cells[i + 1, 2] = Vehicle_no;
                    excel.Cells[i + 1, 3] = Valid_from;
                    excel.Cells[i + 1, 4] = Valid_To;
                    excel.Cells[i + 1, 5] = Holder_type;
                    excel.Cells[i + 1, 6] = Valid_status;
                    i++;
                }


                Microsoft.Office.Interop.Excel.Range Total_range = excel.get_Range("A1", $"F{i + 1}");//
                Total_range.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                Total_range.EntireColumn.AutoFit();

                excel.DisplayAlerts = false;
                string path = $"{Application.StartupPath}\\Holder\\";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                path = path + kv.Key;
                //workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\" + batchlist[kv.Key] + "\\" + kv.Key + "_" + sat.ToString("yyyyMM") + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel5, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workbook.SaveAs(path, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value);
                workbook.Save();
                LogClass.WriteLog($"{kv.Key} Export excel completed!");
                goto Close;
                Close:
                excel.Workbooks.Close();
                excel.Quit();
                GC.Collect();
                Kill(excel);
            }

            Application.Exit();
        }
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
        {
            IntPtr t = new IntPtr(excel.Hwnd);   //得到这个句柄，具体作用是得到这块内存入口 

            int k = 0;
            GetWindowThreadProcessId(t, out k);   //得到本进程唯一标志k
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用
            p.Kill();     //关闭进程k
        }
        #endregion



    }
}
