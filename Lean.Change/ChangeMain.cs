using System;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Lean.Change
{
    public partial class ChangeMain : Form
    {
        public ChangeMain()
        {
            InitializeComponent();
        }

        [DllImport("kernel32.dll")]
        private static extern uint GetTickCount();

        //延时函数
        private static void Delay(uint ms)
        {
            uint start = GetTickCount();
            while (GetTickCount() - start < ms)
            {
                Application.DoEvents();
            }
        }

        private string strMailecn, SerialN, SerialO, SerialS, StN, StC, DestN, EcNo, Updatem, Updatep;
        private string Bomitem;
        private string BomSitem;
        private string BomOitem;
        private string BomNitem;
        private string dbFileName;//下载文件名
        private string bkFileName;//备份文件名
        private string bkPath;//备份位置
        private string SavePath;//获取启动了应用程序的可执行文件的路径，不包括可执行文件的名称
        private string datestr, TotalTime;

        private void ChangeMain_Load(object sender, EventArgs e)
        {
            //txtReaderMaterialText();
            //txtReaderecn();
            //txtReaderecnsub();
            //txtReaderOrder();
            //txtReaderOrdersn();

            //txtReaderTimes();
            //txtReaderMaterial();
            //AccdbReaderitem();
            //AccdbReaderec();
            //txtReaderMaterials();
            //string pathnames=@Application.StartupPath + "\\BackFile\\JA-OS-unicode.txt";
            ////string content = File.ReadAllText(pathnames, Encoding.Default);
            //MessageBox.Show(GetFileEncoding(pathnames).ToString());
            //string content = File.ReadAllText("20190528_ec.txt", Encoding.Default);
            //File.WriteAllText("20190528_ecA.txt", content, Encoding.GetEncoding("GB2312"));

            //Helper_File.GetEncoding("20190528_ec.txt");
            //sss();
            //dattostring();
            //txtReader();

            //延时20分钟'1200000
            //Delay(600000);
            //DownloadExcel();
            //DownloadText();

            DownloadAccess();
            this.Close();
        }

        private string XmlFile, mailTable, logStr;
        public static int rows, mailrows;

        private string connstrfs3 = ConfigurationManager.ConnectionStrings["UploadFs3Serv"].ConnectionString;

        private void DownloadAccess()
        {
            try
            {
                //基本设置
                string[] dbTextList = { "ec.accdb", "ec_sub.accdb", "order.accdb", "serial.accdb", "st.accdb", "item.accdb", "model.accdb" };
                //string[] dbFileList = { "SapcooisData.xlsx", "SapzpbldData.xlsx", "SapzpabdData.xlsx", "Sapser05Data.xlsx", "Sapzc1adData.xlsx", "SapzpabdsubData.xlsx" };
                //string[] dbFileListtext = { "ec.txt" };
                for (int i = 0; i < dbTextList.Count(); i++)
                {
                    datestr = DateTime.Now.ToString("yyyyMMdd");
                    dbFileName = dbTextList[i].ToString();
                    bkFileName = (datestr + '_' + dbTextList[i].ToString()).Trim();
                    SavePath = Application.StartupPath + "\\DownFile\\"; //获取启动了应用程序的可执行文件的路径，不包括可执行文件的名称
                    XmlFile = Application.StartupPath + "\\ChangeSetting.xml";
                    string uid = Helper_Xml.Read(XmlFile, "/Root/FTP/UserInfo", "Uid").ToString();
                    string pwd = Helper_Xml.Read(XmlFile, "/Root/FTP/UserInfo", "Pwd").ToString();
                    string ftpip = Helper_Xml.Read(XmlFile, "/Root/FTP/ServInfo ", "ServIpaddr").ToString();
                    string ftpname = Helper_Xml.Read(XmlFile, "/Root/FTP/ServInfo ", "ServName").ToString();
                    string path = @"ftp://" + ftpip + "/";    //目标路径
                    string pathname = @"ftp://" + ftpname + "/";    //目标路径
                    bkPath = Application.StartupPath + "\\BackFile\\";

                    Helper_Ftp m_ftp = new Helper_Ftp();
                    string[] ftpFileList = Helper_Ftp.GetFileList("/");
                    if (ftpFileList != null)
                    {
                        for (int f = 0; f < ftpFileList.Count(); f++)
                        {
                            string DownFileName = ftpFileList[f].ToString();
                            if (DownFileName == dbFileName)
                            {
                                //下载文件
                                m_ftp.Download(dbFileName);

                                string uptime = DateTime.Now.ToString();
                                TimeSpan ts1 = new TimeSpan(DateTime.Now.Ticks); //获取当前时间的刻度数

                                //ExceltoMssql();
                                AccesstoMsql();

                                //下载完成后删除FTP文件
                                Helper_Ftp.DeleteFileName(dbFileName);
                                //移动前删除已存在的文件
                                Helper_File.DeleteFile(bkPath + bkFileName);
                                //移动文件
                                Helper_File.MoveFile(SavePath + dbFileName, bkPath + bkFileName);

                                //显示运行时间
                                TimeSpan ts2 = new TimeSpan(DateTime.Now.Ticks);
                                TimeSpan ts = ts2.Subtract(ts1).Duration(); //时间差的绝对值
                                string spanTotalSeconds = ts.TotalSeconds.ToString(); //执行时间的总秒数
                                string spanTime = "FTP Uplaod,共花费" + ts.Hours.ToString() + "小时" + ts.Minutes.ToString() + "分" + ts.Seconds.ToString() + "秒！"; //以X小时X分X秒的格式现实执行时间
                                TotalTime = ts.Hours.ToString() + "小时" + ts.Minutes.ToString() + "分" + ts.Seconds.ToString() + "秒！";

                                mailTable = dbFileName + "(" + mailrows + "件)" + "\r\n";
                                string logrows = mailrows + "件,用时：";
                                //添加日志
                                logStr = "INSERT INTO[dbo].[ProcessingLogs] VALUES('" + Guid.NewGuid() + "','Upload','" + dbFileName + "','" + logrows + "'+'" + TotalTime + "','" + "Auto" + "','" + "admin" + "','" + Helper_Hard.GetIPAddress() + "','" + DateTime.Now + "')";

                                Helper_File.SendMail(uptime, mailTable);
                                Helper_Sql Helper_Sql = new Helper_Sql();
                                Helper_Sql.ExecuteNonQuery(logStr);
                            }
                        }
                    }
                }
            }
            catch (ArithmeticException e)
            {
                MessageBox.Show("ArithmeticException Handler: {0}", e.ToString());
            }
            catch (Exception e)
            {
                MessageBox.Show("错误信息如下\r\n  " + e.ToString());
                // 错误处理代码
            }
        }

        [DllImport("kernel32.dll")]
        public static extern bool SetProcessWorkingSetSize(IntPtr process, int minSize, int maxSize);

        public static void GarbageCollect()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public static void FlushMemory()
        {
            GarbageCollect();

            if (Environment.OSVersion.Platform == PlatformID.Win32NT)
            {
                SetProcessWorkingSetSize(System.Diagnostics.Process.GetCurrentProcess().Handle, -1, -1);
            }
        }

        private void AccesstoMsql()
        {
            #region ecn.txt

            if (dbFileName == "ec.accdb")
            {
                string str = "\'";
                //string ss=Helper_Accdb.dataSet("select * from [ec] ", Application.StartupPath + "\\ec.accdb").Tables[0].Rows[0][1].ToString();

                //获取ACCESS表名
                Helper_Accdb.tableName(SavePath + dbFileName);

                // 遍历所有的表
                string sName = "";
                for (int i = 0, maxI = Helper_Accdb.tableName(SavePath + dbFileName).Rows.Count; i < maxI; i++)
                {
                    // 获取第i个Access数据库中的表名
                    string sTempTableName = Helper_Accdb.tableName(SavePath + dbFileName).Rows[i]["TABLE_NAME"].ToString();
                    sName += string.Format("[{0}]\n", sTempTableName);
                }
                DataSet dds = Helper_Accdb.dataSet("SELECT * FROM " + sName + " ; ", SavePath + dbFileName);

                mailrows = dds.Tables[0].Rows.Count;
                rows = dds.Tables[0].Rows.Count;
                //遍历一个表多行多列
                for (int i = 0; i < rows; i++)
                {
                    int comms = dds.Tables[0].Columns.Count;
                    EcNo = dds.Tables[0].Rows[i][0].ToString();
                    strMailecn += EcNo + ";";
                    Helper_Sql Helper_Sql = new Helper_Sql();
                    string non = "SELECT * FROM [dbo].[PP_SapEcn] WHERE D_SAP_ZPABD_Z001= '" + EcNo + "'";
                    if (Helper_Sql.SqlServerRecordCount(non) == 0)
                    {
                        for (int f = 0; f < comms; f++)
                        {
                            if (dds.Tables[0].Rows[i][0].ToString() != "")

                            {
                                str = str + dds.Tables[0].Rows[i][f].ToString().Replace("\'", ",") + "\',\'";
                            }
                        }

                        str = "'" + Guid.NewGuid() + "'," + str + "App自动添加'," + 0 + ",'admin','" + DateTime.Now + "','admin','" + DateTime.Now + "'";

                        string Insql = "INSERT INTO [dbo].[PP_SapEcn] VALUES (" + str + ");";
                        Helper_Sql.ExecuteNonQuery(Insql);
                        //发邮件
                        string APP_Path = Application.StartupPath;//获取启动了应用程序的可执行文件的路径，不包括可执行文件的名称
                        string XmlFile = APP_Path + "\\ChangeMailto.xml";

                        string strMailto = Helper_Xml.Read(XmlFile, "/Root/SendtoList/ENG", "Mail").ToString();
                        //Helper_File.ecnSendMail(DateTime.Now.ToString(), strMailecn);
                        string mailTitle = "设变发行：" + strMailecn;
                        string mailBody = "Dear All,\r\n" + "\r\n" + "此设变技术部门未处理。\r\n" + "请贵部门担当者及时处理为盼。\r\n" + "\r\n" + "よろしくお願いいたします。\r\n" + "\r\n" + "\r\n" + "「设变DB系统\r\n" + DateTime.Now.ToString() + "」\r\n" + "このメッセージはDBシステムから自動で送信されている。\r\n\n";  //发送邮件的正文

                        Helper_File.SendEmail(strMailto, mailTitle, mailBody);
                    }
                    else
                    {
                        string Insql = "Update[dbo].[PP_SapEcn] set Modifier='admin', ModifyTime='" + DateTime.Now + "'where D_SAP_ZPABD_Z001='" + EcNo + "'";
                        Helper_Sql.ExecuteNonQuery(Insql);
                    }

                    str = "\'";
                }
            }

            #endregion ecn.txt

            #region ecnsub.txt

            if (dbFileName == "ec_sub.accdb")
            {
                string str = "\'";
                //string ss=Helper_Accdb.dataSet("select * from [ec] ", Application.StartupPath + "\\ec.accdb").Tables[0].Rows[0][1].ToString();

                //获取ACCESS表名
                Helper_Accdb.tableName(SavePath + dbFileName);

                // 遍历所有的表
                string sName = "";
                for (int i = 0, maxI = Helper_Accdb.tableName(SavePath + dbFileName).Rows.Count; i < maxI; i++)
                {
                    // 获取第i个Access数据库中的表名
                    string sTempTableName = Helper_Accdb.tableName(SavePath + dbFileName).Rows[i]["TABLE_NAME"].ToString();
                    sName += string.Format("[{0}]\n", sTempTableName);
                }
                DataSet dds = Helper_Accdb.dataSet("SELECT * FROM " + sName + "; ", SavePath + dbFileName);
                //遍历一个表多行多列

                mailrows = dds.Tables[0].Rows.Count;
                rows = dds.Tables[0].Rows.Count;

                for (int i = 0; i < rows; i++)
                {
                    int comms = dds.Tables[0].Columns.Count;

                    EcNo = dds.Tables[0].Rows[i][0].ToString();
                    string non = "SELECT * FROM [dbo].[PP_SapEcnSub] WHERE  D_SAP_ZPABD_S001='" + EcNo + "' and D_SAP_ZPABD_S002='" + Bomitem + "' and D_SAP_ZPABD_S003='" + BomSitem + "' and  D_SAP_ZPABD_S004='" + BomOitem + "'  and D_SAP_ZPABD_S008= '" + BomNitem + "'";
                    Helper_Sql Helper_Sql = new Helper_Sql();
                    if (Helper_Sql.SqlServerRecordCount(non) == 0)
                    {
                        if (EcNo != "")
                        {
                            if (Regex.IsMatch(dds.Tables[0].Rows[i][1].ToString().TrimEnd(), @"^[+-]?\d*[.]?\d*$"))
                            {
                                if (dds.Tables[0].Rows[i][1].ToString().TrimEnd() != "")
                                {
                                    Bomitem = dds.Tables[0].Rows[i][1].ToString().TrimEnd();
                                }
                                else
                                {
                                    Bomitem = dds.Tables[0].Rows[i][1].ToString().TrimEnd();
                                }
                            }
                            else
                            {
                                Bomitem = dds.Tables[0].Rows[i][1].ToString().TrimEnd();
                            }

                            if (Regex.IsMatch(dds.Tables[0].Rows[i][2].ToString().TrimEnd(), @"^[+-]?\d*[.]?\d*$"))
                            {
                                if (dds.Tables[0].Rows[i][2].ToString().TrimEnd() != "")
                                {
                                    BomSitem = dds.Tables[0].Rows[i][2].ToString().TrimEnd();
                                }
                                else
                                {
                                    BomSitem = dds.Tables[0].Rows[i][2].ToString().TrimEnd();
                                }
                            }
                            else
                            {
                                BomSitem = dds.Tables[0].Rows[i][2].ToString().TrimEnd();
                            }
                            if (Regex.IsMatch(dds.Tables[0].Rows[i][3].ToString().TrimEnd(), @"^[+-]?\d*[.]?\d*$"))
                            {
                                if (dds.Tables[0].Rows[i][3].ToString().TrimEnd() != "")
                                {
                                    BomOitem = dds.Tables[0].Rows[i][3].ToString().TrimEnd();
                                }
                                else
                                {
                                    BomOitem = dds.Tables[0].Rows[i][3].ToString().TrimEnd();
                                }
                            }
                            else
                            {
                                BomOitem = dds.Tables[0].Rows[i][3].ToString().TrimEnd();
                            }
                            if (Regex.IsMatch(dds.Tables[0].Rows[i][7].ToString().TrimEnd(), @"^[+-]?\d*[.]?\d*$"))
                            {
                                if (dds.Tables[0].Rows[i][7].ToString().TrimEnd() != "")
                                {
                                    BomNitem = dds.Tables[0].Rows[i][7].ToString().TrimEnd();
                                }
                                else
                                {
                                    BomNitem = dds.Tables[0].Rows[i][7].ToString().TrimEnd();
                                }
                            }
                            else
                            {
                                BomNitem = dds.Tables[0].Rows[i][7].ToString().TrimEnd();
                            }
                            for (int f = 0; f < comms; f++)
                            {
                                if (f == 1 || f == 2 || f == 3 || f == 7)
                                {
                                    //string stf = Int64.Parse("00000000005781702606").ToString();
                                    if (Regex.IsMatch(dds.Tables[0].Rows[i][f].ToString().TrimEnd(), @"^/d*[.]?/d*$"))
                                    {
                                        if (dds.Tables[0].Rows[i][f].ToString().TrimEnd() != "")
                                        {
                                            //string ss = dds.Tables[0].Rows[i][f].ToString();
                                            str = str + Int64.Parse(dds.Tables[0].Rows[i][f].ToString()).ToString() + "\',\'";
                                        }
                                        else
                                        {
                                            str = str + dds.Tables[0].Rows[i][f].ToString() + "\',\'";
                                        }
                                    }
                                    else
                                    {
                                        str = str + dds.Tables[0].Rows[i][f].ToString() + "\',\'";
                                    }
                                }
                                else
                                {
                                    str = str + dds.Tables[0].Rows[i][f].ToString() + "\',\'";
                                }
                            }
                            str = "'" + Guid.NewGuid() + "'," + str + "App自动添加'," + 0 + ",'admin','" + DateTime.Now + "','admin','" + DateTime.Now + "'";

                            string Insql = "INSERT INTO [dbo].[PP_SapEcnSub] VALUES (" + str + ");";
                            Helper_Sql.ExecuteNonQuery(Insql);
                        }

                        str = "\'";
                        EcNo = "";
                        Bomitem = "";
                        BomSitem = "";
                        BomOitem = "";
                        BomNitem = "";
                    }
                    else
                    {
                        string Insql = "Update[dbo].[PP_SapEcnSub] set Modifier='admin', ModifyTime='" + DateTime.Now + "'where D_SAP_ZPABD_S001='" + EcNo + "'";
                        Helper_Sql.ExecuteNonQuery(Insql);
                    }
                }

                //foreach (DataRow mDr in Helper_Accdb.dataSet("SELECT ec.[設変番号], ec.[機種], ec.[タイトル], ec.[ステータス], ec.[発行日], ec.[担当], ec.[依頼元], ec.[設変会議], ec.[PP番号], ec.[技術連絡書], ec.[実施], ec.[変更理由(主)], ec.[変更理由(従)], ec.[安全規格], ec.[進行状況], ec.[客先承認], ec.[機番管理], ec.[Sマニュアル訂正], ec.[取説訂正], ec.[カタログ訂正], ec.[作業標準訂正], ec.[技術情報発行], ec.[コスト変動], ec.[単位コスト], ec.[金型改造費], ec.[関連設変], ec.[設変記事] FROM ec; ", Application.StartupPath + "\\ec.accdb").Tables[0].Rows)
                //{
                //    foreach (DataColumn mDc in Helper_Accdb.dataSet("SELECT ec.[設変番号], ec.[機種], ec.[タイトル], ec.[ステータス], ec.[発行日], ec.[担当], ec.[依頼元], ec.[設変会議], ec.[PP番号], ec.[技術連絡書], ec.[実施], ec.[変更理由(主)], ec.[変更理由(従)], ec.[安全規格], ec.[進行状況], ec.[客先承認], ec.[機番管理], ec.[Sマニュアル訂正], ec.[取説訂正], ec.[カタログ訂正], ec.[作業標準訂正], ec.[技術情報発行], ec.[コスト変動], ec.[単位コスト], ec.[金型改造費], ec.[関連設変], ec.[設変記事] FROM ec; ", Application.StartupPath + "\\ec.accdb").Tables[0].Columns)
                //    {
                //        str = mDr[mDc].ToString() + "\',\'";
                //        //Console.WriteLine(mDr[mDc].ToString());
                //    }

                //}
            }

            #endregion ecnsub.txt

            #region item

            if (dbFileName == "item.accdb")
            {
                Helper_Sql Helper_Sql = new Helper_Sql();
                string str = "\'";
                //string ss=Helper_Accdb.dataSet("select * from [ec] ", Application.StartupPath + "\\ec.accdb").Tables[0].Rows[0][1].ToString();

                //获取ACCESS表名
                Helper_Accdb.tableName(SavePath + dbFileName);

                // 遍历所有的表
                string sName = "";
                for (int i = 0, maxI = Helper_Accdb.tableName(SavePath + dbFileName).Rows.Count; i < maxI; i++)
                {
                    // 获取第i个Access数据库中的表名
                    string sTempTableName = Helper_Accdb.tableName(SavePath + dbFileName).Rows[i]["TABLE_NAME"].ToString();
                    sName += string.Format("[{0}]\n", sTempTableName);
                }
                //判断在库检查
                //DataSet dds = Helper_Accdb.dataSet("SELECT * FROM " + sName + " WHERE Plant='H100' AND 検査中在庫登録='X'; ", SavePath  + dbFileName);
                ////DataSet dds = Helper_Accdb.dataSet("SELECT * FROM " + sName + " WHERE (((PRD_100_ZS0PR074DO101.Plant)='C100')); ", SavePath  + dbFileName);
                ////遍历一个表多行多列

                //rows = dds.Tables[0].Rows.Count;

                //for (int i = 0; i < rows; i++)
                //{
                //    ItemNo = dds.Tables[0].Rows[i][2].ToString();
                //    itemx = dds.Tables[0].Rows[i][19].ToString();
                //    Helper_Sql Helper_Sql = new Helper_Sql();
                //    string Insql = "UPDATE [dbo].[PP_SapMaterial] SET [D_SAP_ZCA1D_Z019]='" + itemx + "' WHERE  D_SAP_ZCA1D_Z002='" + ItemNo + "'; ";
                //    Helper_Sql.ExecuteNonQuery(Insql);

                //    str = "\'";
                //    ItemNo = "";
                //}
                //更新Item
                DataSet ddsC100 = Helper_Accdb.dataSet("SELECT * FROM " + sName + "; ", SavePath + dbFileName);

                //DataTable ddtC100 = Helper_Accdb.dataTable("SELECT * FROM " + sName + " WHERE Plant='C100'; ", SavePath + dbFileName);
                //Helper_Sql.GetTablesKeys("D_SAP_ZCA1D_Z002");
                //Helper_Sql.UpdateExistData(ddtC100, "PP_SapMaterial");

                mailrows = ddsC100.Tables[0].Rows.Count;
                rows = ddsC100.Tables[0].Rows.Count;
                string non = "DELETE  [dbo].[PP_SapMaterial_TEMP] ;";

                Helper_Sql.ExecuteNonQuery(non);
                for (int fi = 0; fi < rows; fi++)
                {
                    //plnt = ddsC100.Tables[0].Rows[fi][1].ToString();
                    //Pper = ddsC100.Tables[0].Rows[fi][9].ToString();
                    //Ppur = ddsC100.Tables[0].Rows[fi][10].ToString();
                    //Pgroup = ddsC100.Tables[0].Rows[fi][11].ToString();
                    //Pblock = ddsC100.Tables[0].Rows[fi][12].ToString();
                    //Pavgprice = ddsC100.Tables[0].Rows[fi][26].ToString();
                    //itemx = ddsC100.Tables[0].Rows[0][19].ToString();
                    //Peol = ddsC100.Tables[0].Rows[fi][34].ToString();
                    //ItemNo = ddsC100.Tables[0].Rows[fi][2].ToString();
                    //UpdateInv = ddsC100.Tables[0].Rows[fi][33].ToString();
                    //pStoc = ddsC100.Tables[0].Rows[fi][30].ToString();

                    int comms = ddsC100.Tables[0].Columns.Count;

                    for (int f = 1; f < comms; f++)
                    {
                        str = str + ddsC100.Tables[0].Rows[fi][f].ToString().Replace("\'", ",") + "\',\'";
                    }
                    //str = "'" + Guid.NewGuid() + "'," + str + "','App自动添加'," + 0 + ",'admin','" + DateTime.Now + "','',''";
                    str = "'" + Guid.NewGuid() + "'," + str + "App自动添加'," + 0 + ",'admin','" + DateTime.Now + "','admin','" + DateTime.Now + "'";

                    string Insql = "INSERT INTO [dbo].[PP_SapMaterial_TEMP] VALUES (" + str + ");";
                    Helper_Sql.ExecuteNonQuery(Insql);

                    str = "\'";
                    //ItemNo = "";
                }

                //foreach (DataRow mDr in Helper_Accdb.dataSet("SELECT ec.[設変番号], ec.[機種], ec.[タイトル], ec.[ステータス], ec.[発行日], ec.[担当], ec.[依頼元], ec.[設変会議], ec.[PP番号], ec.[技術連絡書], ec.[実施], ec.[変更理由(主)], ec.[変更理由(従)], ec.[安全規格], ec.[進行状況], ec.[客先承認], ec.[機番管理], ec.[Sマニュアル訂正], ec.[取説訂正], ec.[カタログ訂正], ec.[作業標準訂正], ec.[技術情報発行], ec.[コスト変動], ec.[単位コスト], ec.[金型改造費], ec.[関連設変], ec.[設変記事] FROM ec; ", Application.StartupPath + "\\ec.accdb").Tables[0].Rows)
                //{
                //    foreach (DataColumn mDc in Helper_Accdb.dataSet("SELECT ec.[設変番号], ec.[機種], ec.[タイトル], ec.[ステータス], ec.[発行日], ec.[担当], ec.[依頼元], ec.[設変会議], ec.[PP番号], ec.[技術連絡書], ec.[実施], ec.[変更理由(主)], ec.[変更理由(従)], ec.[安全規格], ec.[進行状況], ec.[客先承認], ec.[機番管理], ec.[Sマニュアル訂正], ec.[取説訂正], ec.[カタログ訂正], ec.[作業標準訂正], ec.[技術情報発行], ec.[コスト変動], ec.[単位コスト], ec.[金型改造費], ec.[関連設変], ec.[設変記事] FROM ec; ", Application.StartupPath + "\\ec.accdb").Tables[0].Columns)
                //    {
                //        str = mDr[mDc].ToString() + "\',\'";
                //        //Console.WriteLine(mDr[mDc].ToString());
                //    }

                //}
            }

            #endregion item

            #region itemUpdate&Insert

            if (dbFileName == "item.accdb")
            {
                string str = "\'";
                Helper_Sql Helper_Sql = new Helper_Sql();
                string Updatesql = "UPDATE  PP_SapMaterial " +
                                    "SET D_SAP_ZCA1D_Z004 = PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z004,D_SAP_ZCA1D_Z006 = PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z006,D_SAP_ZCA1D_Z009 = PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z009," +
                                    "D_SAP_ZCA1D_Z010 = PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z010, " +
                                    "D_SAP_ZCA1D_Z013 = PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z013,D_SAP_ZCA1D_Z015 = PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z015,D_SAP_ZCA1D_Z017 = PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z017," +
                                    "D_SAP_ZCA1D_Z019 = PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z019," +

                                    "D_SAP_ZCA1D_Z026 = PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z026, D_SAP_ZCA1D_Z030 = PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z030," +
                                    "D_SAP_ZCA1D_Z031 = PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z031" +
                                    ", D_SAP_ZCA1D_Z033 = PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z033, D_SAP_ZCA1D_Z034 = PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z034" +
                                    ", Modifier = 'Day_Update', ModifyTime = (select CONVERT(varchar, GETDATE(), 120))" +
                                    " FROM PP_SapMaterial inner join PP_SapMaterial_TEMP" +
                                    " ON PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z002 = PP_SapMaterial.D_SAP_ZCA1D_Z002" +
                                    " where PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z001 = 'C100';" +
                                    " UPDATE PP_SapMaterial" +
                                    " SET D_SAP_ZCA1D_Z009 = PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z009" +
                                    " , Modifier = 'Day_Update', ModifyTime = (select CONVERT(varchar, GETDATE(), 120))" +
                                    " FROM PP_SapMaterial inner join PP_SapMaterial_TEMP" +
                                    " ON PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z002 = PP_SapMaterial.D_SAP_ZCA1D_Z002" +
                                    " where PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z001 = 'H100'; ";

                Helper_Sql.ExecuteNonQuery(Updatesql);

                string InsertSql = "  SELECT NEWID()GUID, PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z001,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z002,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z003,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z004,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z005,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z006 " +
                                    ", PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z007,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z008,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z009,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z010,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z011,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z012,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z013 " +
                                    ",PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z014,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z015,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z016,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z017,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z018,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z019,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z020 " +
                                    ",PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z021,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z022,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z023,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z024,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z025,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z026,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z027 " +
                                    ",PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z028,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z029,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z030,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z031,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z032,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z033,PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z034 FROM PP_SapMaterial_TEMP LEFT JOIN [PP_SapMaterial] ON  " +
                                    "  PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z002 =[PP_SapMaterial].D_SAP_ZCA1D_Z002 WHERE " +
                                    "  PP_SapMaterial_TEMP.D_SAP_ZCA1D_Z001 = 'C100' AND " +
                                    "    [PP_SapMaterial].D_SAP_ZCA1D_Z002 IS NULL";

                if (Helper_Sql.SqlServerRecordCount(InsertSql) > 0)
                {
                    DataSet DSItem = Helper_Sql.GetDataSetValue(InsertSql, "Item");

                    mailrows = DSItem.Tables["Item"].Rows.Count;
                    rows = DSItem.Tables["Item"].Rows.Count;
                    for (int fi = 0; fi < rows; fi++)
                    {
                        int comms = DSItem.Tables["Item"].Columns.Count;

                        for (int f = 1; f < comms; f++)
                        {
                            str = str + DSItem.Tables["Item"].Rows[fi][f].ToString().Replace("\'", ",") + "\',\'";
                        }
                        //str = "'" + Guid.NewGuid() + "'," + str + "','App自动添加'," + 0 + ",'admin','" + DateTime.Now + "','',''";
                        str = "'" + Guid.NewGuid() + "'," + str + "App自动添加'," + 0 + ",'admin','" + DateTime.Now + "','admin','" + DateTime.Now + "'";

                        string Insql = "INSERT INTO [dbo].[PP_SapMaterial] VALUES (" + str + ");";
                        Helper_Sql.ExecuteNonQuery(Insql);

                        str = "\'";
                    }
                }
            }

            #endregion itemUpdate&Insert

            #region order.txt

            if (dbFileName == "order.accdb")
            {
                string str = "\'";
                Helper_Sql Helper_Sql = new Helper_Sql();
                //string ss=Helper_Accdb.dataSet("select * from [ec] ", Application.StartupPath + "\\ec.accdb").Tables[0].Rows[0][1].ToString();

                //获取ACCESS表名
                Helper_Accdb.tableName(SavePath + dbFileName);

                // 遍历所有的表
                string sName = "";
                for (int i = 0, maxI = Helper_Accdb.tableName(SavePath + dbFileName).Rows.Count; i < maxI; i++)
                {
                    // 获取第i个Access数据库中的表名
                    string sTempTableName = Helper_Accdb.tableName(SavePath + dbFileName).Rows[i]["TABLE_NAME"].ToString();
                    sName += string.Format("[{0}];\n", sTempTableName);
                }
                DataSet dds = Helper_Accdb.dataSet("SELECT * FROM " + sName, SavePath + dbFileName);
                //遍历一个表多行多列

                mailrows = dds.Tables[0].Rows.Count;
                rows = dds.Tables[0].Rows.Count;
                string non = "DELETE  [dbo].[PP_SapOrders_TEMP] ;";
                Helper_Sql.ExecuteNonQuery(non);
                for (int i = 0; i < rows; i++)
                {
                    //OrderNo = dds.Tables[0].Rows[i][1].ToString();
                    //string olot= dds.Tables[0].Rows[i][3].ToString();
                    //string odate= dds.Tables[0].Rows[i][6].ToString();
                    //string oss= dds.Tables[0].Rows[i][7].ToString();
                    //oqty = dds.Tables[0].Rows[i][4].ToString();
                    //nqty = dds.Tables[0].Rows[i][5].ToString();

                    int comms = dds.Tables[0].Columns.Count;

                    for (int f = 0; f < comms; f++)
                    {
                        str = str + dds.Tables[0].Rows[i][f].ToString() + "\',\'";
                    }
                    str = "'" + Guid.NewGuid() + "'," + str + "App自动添加'," + 0 + ",'admin','" + DateTime.Now + "','admin','" + DateTime.Now + "'";

                    string Insql = "INSERT INTO [dbo].[PP_SapOrders_TEMP] VALUES (" + str + ");";
                    Helper_Sql.ExecuteNonQuery(Insql);

                    str = "\'";
                }
                //存在就更新
                string UpdateSql = "  UPDATE  [PP_SapOrders] " +
                                    " SET[D_SAP_COOIS_C004] =[PP_SapOrders_TEMP].[D_SAP_COOIS_C004], " +
                                    " [D_SAP_COOIS_C005]=[PP_SapOrders_TEMP].[D_SAP_COOIS_C005], " +
                                    " [D_SAP_COOIS_C006]=[PP_SapOrders_TEMP].[D_SAP_COOIS_C006], " +
                                    " [D_SAP_COOIS_C007]=[PP_SapOrders_TEMP].[D_SAP_COOIS_C007], " +
                                    " [D_SAP_COOIS_C008]=[PP_SapOrders_TEMP].[D_SAP_COOIS_C008] " +
                                    " ,Modifier ='Day_Update',ModifyTime=(select CONVERT(varchar, GETDATE(),120)) " +
                                    " FROM[PP_SapOrders] inner join[PP_SapOrders_TEMP] " +
                                    " ON[PP_SapOrders_TEMP].[D_SAP_COOIS_C002]=[PP_SapOrders].[D_SAP_COOIS_C002] " +
                                    " WHERE [PP_SapOrders_TEMP].[D_SAP_COOIS_C001]='C100'";
                Helper_Sql.ExecuteNonQuery(UpdateSql);

                //不存在就新增
                string InsertSql = "SELECT NEWID() [GUID] " +
                                    ",[PP_SapOrders_TEMP].[D_SAP_COOIS_C001],[PP_SapOrders_TEMP].[D_SAP_COOIS_C002] " +
                                    ",[PP_SapOrders_TEMP].[D_SAP_COOIS_C003],[PP_SapOrders_TEMP].[D_SAP_COOIS_C004] " +
                                    ",[PP_SapOrders_TEMP].[D_SAP_COOIS_C005],[PP_SapOrders_TEMP].[D_SAP_COOIS_C006] " +
                                    ",[PP_SapOrders_TEMP].[D_SAP_COOIS_C007],[PP_SapOrders_TEMP].[D_SAP_COOIS_C008] " +
                                    " ,[PP_SapOrders_TEMP].[D_SAP_COOIS_C009]       FROM[Sap_Data].[dbo].[PP_SapOrders_TEMP] " +
                                    "        LEFT JOIN[Sap_Data].[dbo].[PP_SapOrders] " +
                                    "        ON [PP_SapOrders_TEMP].[D_SAP_COOIS_C002]=[PP_SapOrders].[D_SAP_COOIS_C002] " +
                                    "        WHERE [PP_SapOrders_TEMP].[D_SAP_COOIS_C001]='C100' AND[PP_SapOrders].[D_SAP_COOIS_C002] IS NULL";
                DataSet DSItem = Helper_Sql.GetDataSetValue(InsertSql, "Item");

                int Insrows = DSItem.Tables["Item"].Rows.Count;
                int Inscols = DSItem.Tables["Item"].Columns.Count;
                for (int fi = 0; fi < Insrows; fi++)
                {
                    for (int f = 1; f < Inscols; f++)
                    {
                        str = str + DSItem.Tables["Item"].Rows[fi][f].ToString() + "\',\'";
                    }
                    str = "'" + Guid.NewGuid() + "'," + str + "App自动添加'," + 0 + ",'admin','" + DateTime.Now + "','admin','" + DateTime.Now + "'";
                    string Insqls = "INSERT INTO [dbo].[PP_SapOrders] VALUES (" + str + ");";
                    Helper_Sql.ExecuteNonQuery(Insqls);

                    str = "\'";
                }
            }

            #endregion order.txt

            #region ordersn.txt

            if (dbFileName == "serial.accdb")
            {
                Helper_Sql Helper_Sql = new Helper_Sql();

                string str = "\'";
                //string ss=Helper_Accdb.dataSet("select * from [ec] ", Application.StartupPath + "\\ec.accdb").Tables[0].Rows[0][1].ToString();

                //获取ACCESS表名
                Helper_Accdb.tableName(SavePath + dbFileName);

                // 遍历所有的表
                string sName = "";
                for (int i = 0, maxI = Helper_Accdb.tableName(SavePath + dbFileName).Rows.Count; i < maxI; i++)
                {
                    // 获取第i个Access数据库中的表名
                    string sTempTableName = Helper_Accdb.tableName(SavePath + dbFileName).Rows[i]["TABLE_NAME"].ToString();
                    sName += string.Format("{0}\n", sTempTableName);
                }
                DataSet dds = Helper_Accdb.dataSet("SELECT * FROM " + sName + "; ", SavePath + dbFileName);
                //遍历一个表多行多列

                mailrows = dds.Tables[0].Rows.Count;
                rows = dds.Tables[0].Rows.Count;

                for (int i = 0; i < rows; i++)
                {
                    int comms = dds.Tables[0].Columns.Count;
                    SerialO = dds.Tables[0].Rows[i][2].ToString();
                    SerialN = dds.Tables[0].Rows[i][3].ToString();
                    SerialS = dds.Tables[0].Rows[i][4].ToString();
                    for (int f = 1; f < comms; f++)
                    {
                        str = str + dds.Tables[0].Rows[i][f].ToString() + "\',\'";
                    }
                    str = "'" + Guid.NewGuid() + "'," + str + "App自动添加'," + 0 + ",'admin','" + DateTime.Now + "','admin','" + DateTime.Now + "'";

                    string non = "SELECT * FROM [dbo].[PP_SapOrderSerial] WHERE D_SAP_SER05_C002='" + SerialO + "' and D_SAP_SER05_C003= '" + SerialN + "' and D_SAP_SER05_C004='" + SerialS + "'";
                    if (Helper_Sql.SqlServerRecordCount(non) == 0)
                    {
                        string Insql = "INSERT INTO [dbo].[PP_SapOrderSerial] VALUES (" + str + ");";
                        Helper_Sql.ExecuteNonQuery(Insql);
                    }
                    //存在就不操作
                    //else
                    //{
                    //    string Upsql = "UPDATE [dbo].[PP_SapOrderSerial] WHERE D_SAP_SER05_C002='" + SerialO + "' and D_SAP_SER05_C003= '" + SerialN + "' and D_SAP_SER05_C004='" + SerialS + "';";
                    //    Helper_Sql.ExecuteNonQuery(Upsql);
                    //}

                    str = "\'";
                    SerialO = "";
                    SerialN = "";
                    SerialS = "";
                }

                //foreach (DataRow mDr in Helper_Accdb.dataSet("SELECT ec.[設変番号], ec.[機種], ec.[タイトル], ec.[ステータス], ec.[発行日], ec.[担当], ec.[依頼元], ec.[設変会議], ec.[PP番号], ec.[技術連絡書], ec.[実施], ec.[変更理由(主)], ec.[変更理由(従)], ec.[安全規格], ec.[進行状況], ec.[客先承認], ec.[機番管理], ec.[Sマニュアル訂正], ec.[取説訂正], ec.[カタログ訂正], ec.[作業標準訂正], ec.[技術情報発行], ec.[コスト変動], ec.[単位コスト], ec.[金型改造費], ec.[関連設変], ec.[設変記事] FROM ec; ", Application.StartupPath + "\\ec.accdb").Tables[0].Rows)
                //{
                //    foreach (DataColumn mDc in Helper_Accdb.dataSet("SELECT ec.[設変番号], ec.[機種], ec.[タイトル], ec.[ステータス], ec.[発行日], ec.[担当], ec.[依頼元], ec.[設変会議], ec.[PP番号], ec.[技術連絡書], ec.[実施], ec.[変更理由(主)], ec.[変更理由(従)], ec.[安全規格], ec.[進行状況], ec.[客先承認], ec.[機番管理], ec.[Sマニュアル訂正], ec.[取説訂正], ec.[カタログ訂正], ec.[作業標準訂正], ec.[技術情報発行], ec.[コスト変動], ec.[単位コスト], ec.[金型改造費], ec.[関連設変], ec.[設変記事] FROM ec; ", Application.StartupPath + "\\ec.accdb").Tables[0].Columns)
                //    {
                //        str = mDr[mDc].ToString() + "\',\'";
                //        //Console.WriteLine(mDr[mDc].ToString());
                //    }

                //}
            }

            #endregion ordersn.txt

            #region st.txt

            if (dbFileName == "st.accdb")
            {
                string str = "\'";
                //string ss=Helper_Accdb.dataSet("select * from [ec] ", Application.StartupPath + "\\ec.accdb").Tables[0].Rows[0][1].ToString();

                //获取ACCESS表名
                Helper_Accdb.tableName(SavePath + dbFileName);

                // 遍历所有的表
                string sName = "";
                for (int i = 0, maxI = Helper_Accdb.tableName(SavePath + dbFileName).Rows.Count; i < maxI; i++)
                {
                    // 获取第i个Access数据库中的表名
                    string sTempTableName = Helper_Accdb.tableName(SavePath + dbFileName).Rows[i]["TABLE_NAME"].ToString();
                    sName += string.Format("{0}\n", sTempTableName);
                }
                DataSet dds = Helper_Accdb.dataSet("SELECT * FROM " + sName + "; ", SavePath + dbFileName);
                //遍历一个表多行多列

                mailrows = dds.Tables[0].Rows.Count;
                rows = dds.Tables[0].Rows.Count;

                for (int i = 0; i < rows; i++)
                {
                    int comms = dds.Tables[0].Columns.Count;
                    StN = dds.Tables[0].Rows[i][1].ToString();
                    StC = dds.Tables[0].Rows[i][2].ToString();
                    Updatep = dds.Tables[0].Rows[i][4].ToString();
                    Updatem = dds.Tables[0].Rows[i][6].ToString();
                    for (int f = 0; f < comms; f++)
                    {
                        str = str + dds.Tables[0].Rows[i][f].ToString() + "\',\'";
                    }
                    str = "'" + Guid.NewGuid() + "'," + str + "App自动添加'," + 0 + ",'admin','" + DateTime.Now + "','admin','" + DateTime.Now + "'";
                    Helper_Sql Helper_Sql = new Helper_Sql();

                    string non = "SELECT * FROM [dbo].[PP_SapManhour] WHERE D_SAP_ZPBLD_Z002= '" + StN + "' and D_SAP_ZPBLD_Z003='" + StC + "'";
                    if (Helper_Sql.SqlServerRecordCount(non) == 0)
                    {
                        string Insql = "INSERT INTO [dbo].[PP_SapManhour] VALUES (" + str + ");";
                        Helper_Sql.ExecuteNonQuery(Insql);
                    }
                    //存在就更新
                    else
                    {
                        string Insql = "UPDATE[dbo].[PP_SapManhour]  SET D_SAP_ZPBLD_Z005='" + Updatem + "',D_SAP_ZPBLD_Z007='" + Updatem + "',Modifier='Day_update',ModifyTime='" + DateTime.Now + "' WHERE D_SAP_ZPBLD_Z002= '" + StN + "' and D_SAP_ZPBLD_Z003='" + StC + "'; ";
                        Helper_Sql.ExecuteNonQuery(Insql);
                    }

                    str = "\'";
                    StN = "";
                    StC = "";
                }

                //foreach (DataRow mDr in Helper_Accdb.dataSet("SELECT ec.[設変番号], ec.[機種], ec.[タイトル], ec.[ステータス], ec.[発行日], ec.[担当], ec.[依頼元], ec.[設変会議], ec.[PP番号], ec.[技術連絡書], ec.[実施], ec.[変更理由(主)], ec.[変更理由(従)], ec.[安全規格], ec.[進行状況], ec.[客先承認], ec.[機番管理], ec.[Sマニュアル訂正], ec.[取説訂正], ec.[カタログ訂正], ec.[作業標準訂正], ec.[技術情報発行], ec.[コスト変動], ec.[単位コスト], ec.[金型改造費], ec.[関連設変], ec.[設変記事] FROM ec; ", Application.StartupPath + "\\ec.accdb").Tables[0].Rows)
                //{
                //    foreach (DataColumn mDc in Helper_Accdb.dataSet("SELECT ec.[設変番号], ec.[機種], ec.[タイトル], ec.[ステータス], ec.[発行日], ec.[担当], ec.[依頼元], ec.[設変会議], ec.[PP番号], ec.[技術連絡書], ec.[実施], ec.[変更理由(主)], ec.[変更理由(従)], ec.[安全規格], ec.[進行状況], ec.[客先承認], ec.[機番管理], ec.[Sマニュアル訂正], ec.[取説訂正], ec.[カタログ訂正], ec.[作業標準訂正], ec.[技術情報発行], ec.[コスト変動], ec.[単位コスト], ec.[金型改造費], ec.[関連設変], ec.[設変記事] FROM ec; ", Application.StartupPath + "\\ec.accdb").Tables[0].Columns)
                //    {
                //        str = mDr[mDc].ToString() + "\',\'";
                //        //Console.WriteLine(mDr[mDc].ToString());
                //    }

                //}
            }

            #endregion st.txt

            #region model.txt

            if (dbFileName == "model.accdb")
            {
                string str = "\'";
                //string ss=Helper_Accdb.dataSet("select * from [ec] ", Application.StartupPath + "\\ec.accdb").Tables[0].Rows[0][1].ToString();

                //获取ACCESS表名
                Helper_Accdb.tableName(SavePath + dbFileName);

                // 遍历所有的表
                string sName = "";
                for (int i = 0, maxI = Helper_Accdb.tableName(SavePath + dbFileName).Rows.Count; i < maxI; i++)
                {
                    // 获取第i个Access数据库中的表名
                    string sTempTableName = Helper_Accdb.tableName(SavePath + dbFileName).Rows[i]["TABLE_NAME"].ToString();
                    sName += string.Format("{0}\n", sTempTableName);
                }
                DataSet dds = Helper_Accdb.dataSet("SELECT * FROM " + sName + "; ", SavePath + dbFileName);
                //遍历一个表多行多列

                mailrows = dds.Tables[0].Rows.Count;
                rows = dds.Tables[0].Rows.Count;

                for (int i = 0; i < rows; i++)
                {
                    int comms = dds.Tables[0].Columns.Count;
                    DestN = dds.Tables[0].Rows[i][1].ToString();
                    for (int f = 1; f < comms; f++)
                    {
                        str = str + dds.Tables[0].Rows[i][f].ToString() + "\',\'";
                    }
                    str = "'" + Guid.NewGuid() + "'," + str + "App自动添加'," + 0 + ",'admin','" + DateTime.Now + "','admin','" + DateTime.Now + "'";
                    Helper_Sql Helper_Sql = new Helper_Sql();

                    string non = "SELECT * FROM [dbo].[PP_SapModelDest] WHERE D_SAP_DEST_Z001= '" + DestN + "'";
                    if (Helper_Sql.SqlServerRecordCount(non) == 0)
                    {
                        string Insql = "INSERT INTO [dbo].[PP_SapModelDest] VALUES (" + str + ");";
                        Helper_Sql.ExecuteNonQuery(Insql);
                    }
                    //存在就不操作
                    str = "\'";
                    DestN = "";
                }

                //foreach (DataRow mDr in Helper_Accdb.dataSet("SELECT ec.[設変番号], ec.[機種], ec.[タイトル], ec.[ステータス], ec.[発行日], ec.[担当], ec.[依頼元], ec.[設変会議], ec.[PP番号], ec.[技術連絡書], ec.[実施], ec.[変更理由(主)], ec.[変更理由(従)], ec.[安全規格], ec.[進行状況], ec.[客先承認], ec.[機番管理], ec.[Sマニュアル訂正], ec.[取説訂正], ec.[カタログ訂正], ec.[作業標準訂正], ec.[技術情報発行], ec.[コスト変動], ec.[単位コスト], ec.[金型改造費], ec.[関連設変], ec.[設変記事] FROM ec; ", Application.StartupPath + "\\ec.accdb").Tables[0].Rows)
                //{
                //    foreach (DataColumn mDc in Helper_Accdb.dataSet("SELECT ec.[設変番号], ec.[機種], ec.[タイトル], ec.[ステータス], ec.[発行日], ec.[担当], ec.[依頼元], ec.[設変会議], ec.[PP番号], ec.[技術連絡書], ec.[実施], ec.[変更理由(主)], ec.[変更理由(従)], ec.[安全規格], ec.[進行状況], ec.[客先承認], ec.[機番管理], ec.[Sマニュアル訂正], ec.[取説訂正], ec.[カタログ訂正], ec.[作業標準訂正], ec.[技術情報発行], ec.[コスト変動], ec.[単位コスト], ec.[金型改造費], ec.[関連設変], ec.[設変記事] FROM ec; ", Application.StartupPath + "\\ec.accdb").Tables[0].Columns)
                //    {
                //        str = mDr[mDc].ToString() + "\',\'";
                //        //Console.WriteLine(mDr[mDc].ToString());
                //    }

                //}
            }

            #endregion model.txt
        }
    }
}