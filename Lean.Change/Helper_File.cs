using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
namespace Lean.Change
{
    class Helper_File
    {
        #region Files类的文件操作方法（创建、复制、删除、移动、追加、打开、设置属性等）
        /// <summary>
        /// 1、创建文件方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void CreateFile(string path)
        {
            //参数1：指定要判断的文件路径
            if (!File.Exists(path))
            {
                //参数1：要创建的文件路径，包含文件名称、后缀等
                FileStream fs = File.Create(path);
                fs.Close();

            }
        }

        /// <summary>
        ///2、 打开文件的方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void OpenFile(string path)
        {
            if (File.Exists(path))
            {
                //参数1：要打开的文件路径，参数2：打开的文件方式
                FileStream fs = File.Open(path, FileMode.Append);
                //字节数组
                byte[] bytes = { (byte)'h', (byte)'e', (byte)'l', (byte)'l', (byte)'o' };
                //通过字符流写入文件
                fs.Write(bytes, 0, bytes.Length);
                fs.Close();

            }
        }

        /// <summary>
        /// 3、追加文件内容方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void AppendFile(string path, string txt)
        {
            string appendtext = txt;
            if (File.Exists(path))
            {
                //参数1：要追加的文件路径，参数2：追加的内容
                File.AppendAllText(path, appendtext);

            }
        }


        /// <summary>
        /// 4、复制文件方法(只能在同个盘符进行操作)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void CopyFile(string oldpath, string newpath)
        {
            if (File.Exists(oldpath))
            {
                //参数1：要复制的源文件路径，参数2：复制后的目标文件路径，参数3：是否覆盖相同文件名
                File.Copy(oldpath, newpath, true);

            }

        }

        /// <summary>
        /// 5、移动文件方法(只能在同个盘符进行操作)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void MoveFile(string oldpath, string newpath)
        {
            if (File.Exists(oldpath))
            {
                //参数1：要移动的源文件路径，参数2：移动后的目标文件路径
                File.Move(oldpath, newpath);

            }
        }

        /// <summary>
        /// 6、删除文件方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void DeleteFile(string path)
        {
            if (File.Exists(path))
            {
                //参数1：要删除的文件路径
                File.Delete(path);

            }
        }

        /// <summary>
        ////7、设置文件属性方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void SetFile(string path)
        {
            if (File.Exists(path))
            {
                //参数1：要设置属性的文件路径，参数2：设置的属性类型（只读、隐藏等）
                File.SetAttributes(path, FileAttributes.Hidden);

            }
        }

        #endregion
        //发送电子邮件成功返回True，失败返回False
        public static bool SendMail(string uptime, string upfile)
        {

            string APP_Path = Application.StartupPath;//获取启动了应用程序的可执行文件的路径，不包括可执行文件的名称            
            string XmlFile = APP_Path + "\\ChangeSetting.xml";


            string mailto = Helper_Xml.Read(XmlFile, "/Root/Send/TO", "Mail").ToString();
            string mailtoname = Helper_Xml.Read(XmlFile, "/Root/Send/TO", "Name").ToString();
            //string mailcc = Helper_Xml.Read(XmlFile, "/Root/Send/CC", "Mail").ToString();
            //string mailccname = Helper_Xml.Read(XmlFile, "/Root/Send/CC", "Name").ToString();
            string mailfrom = Helper_Xml.Read(XmlFile, "/Root/Send/FR", "Mail").ToString();
            string mailfromname = Helper_Xml.Read(XmlFile, "/Root/Send/FR", "Name").ToString();
            string mailsubject = Helper_Xml.Read(XmlFile, "/Root/Send/SU", "Subject").ToString();
            string mailpwd = Helper_Xml.Read(XmlFile, "/Root/Send/FR", "Pwd").ToString();
            //MailAddress from = new MailAddress("cjh@teac.com.cn", "电脑课程建红");
            //收件人地址
            MailAddress to = new MailAddress(mailto, mailtoname);
            //MailAddress cc = new MailAddress(mailcc, mailccname);

            MailMessage message = new MailMessage();

            message.To.Add(to);
            //message.CC.Add(cc);
            message.From = new MailAddress(mailfrom, mailfromname, System.Text.Encoding.UTF8);


            //添加附件，判断文件存在就添加
            //if (System.IO.File.Exists(this.txtAttachment.Text))
            //{
            //    Attachment item = new Attachment(this.txtAttachment.Text, MediaTypeNames.Text.Plain);
            //    message.Attachments.Add(item);
            //}
            message.Subject = mailsubject + ',' + uptime; // 设置邮件的标题
            message.Body = "Dear All,\r\n" + "下記のデータテーブル\r\n" + upfile + "のレコードがデータベースにアップデート完了されました。\r\n" + "ご確認お願い致します。\r\n「" + Helper_Hard.GetComputerName() + "\r\n" + Helper_Hard.GetIPAddress() + "\r\n" + Helper_Hard.GetUserName() + "\r\n" + DateTime.Now.ToString() + "」\r\n" + "このメールはシステムより自動送信されています。\r\nご返信は受付できませんので、ご了承ください。\r\n\n";  //发送邮件的正文
            message.BodyEncoding = System.Text.Encoding.Default;
            //MailAddress other = new MailAddress("davische@teac.com.cn");
            //message.CC.Add(other); //添加抄送人
            //创建一个SmtpClient 类的新实例,并初始化实例的SMTP 事务的服务器
            SmtpClient client = new SmtpClient(@"192.168.16.254");
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.EnableSsl = false;
            //身份认证
            client.Credentials = new System.Net.NetworkCredential(mailfromname, mailpwd);
            bool ret = true; //返回值
            try
            {
                client.Send(message);
            }
            catch (SmtpException ex)
            {
                MessageBox.Show(ex.Message);
                ret = false;
            }
            catch (Exception ex2)
            {
                MessageBox.Show(ex2.Message);
                ret = false;
            }
            return ret;
        }

        public static bool ecnSendMail(string uptime, string ennno)
        {

            string APP_Path = Application.StartupPath;//获取启动了应用程序的可执行文件的路径，不包括可执行文件的名称            
            string XmlFile = APP_Path + "\\ChangeMailto.xml";


            string mailto = Helper_Xml.Read(XmlFile, "/Root/Send/TO", "Mail").ToString();
            string mailtoname = Helper_Xml.Read(XmlFile, "/Root/Send/TO", "Name").ToString();
            string mailcc = Helper_Xml.Read(XmlFile, "/Root/Send/CC", "Mail").ToString();
            string mailccname = Helper_Xml.Read(XmlFile, "/Root/Send/CC", "Name").ToString();
            string mailfrom = Helper_Xml.Read(XmlFile, "/Root/Send/FR", "Mail").ToString();
            string mailfromname = Helper_Xml.Read(XmlFile, "/Root/Send/FR", "Name").ToString();
            string mailsubject = Helper_Xml.Read(XmlFile, "/Root/Send/SU", "Subject").ToString();
            string mailpwd = Helper_Xml.Read(XmlFile, "/Root/Send/FR", "Pwd").ToString();
            //MailAddress from = new MailAddress("cjh@teac.com.cn", "电脑课程建红");
            //收件人地址
            MailAddress to = new MailAddress(mailto, mailtoname);
            MailAddress cc = new MailAddress(mailcc, mailccname);

            MailMessage message = new MailMessage();

            message.To.Add(to);
            message.CC.Add(cc);
            message.From = new MailAddress(mailfrom, mailfromname, System.Text.Encoding.UTF8);


            //添加附件，判断文件存在就添加
            //if (System.IO.File.Exists(this.txtAttachment.Text))
            //{
            //    Attachment item = new Attachment(this.txtAttachment.Text, MediaTypeNames.Text.Plain);
            //    message.Attachments.Add(item);
            //}
            message.Subject = mailsubject + ',' + uptime; // 设置邮件的标题
            message.Body = "Dear All,\r\n" + "请及时处理以下设变：\r\n" + ennno + "\r\n" + "ご確認お願い致します。\r\n此邮件不用回复！\r\n\r\n\r\n「" + Helper_Hard.GetComputerName() + "\r\n" + Helper_Hard.GetIPAddress() + "\r\n" + Helper_Hard.GetUserName() + "\r\n" + DateTime.Now.ToString() + "」\r\n" + "このメールは、システムより自動配信されています。\r\n返信は受付できませんので、ご了承ください。\r\n\n";  //发送邮件的正文
            message.BodyEncoding = System.Text.Encoding.Default;
            //MailAddress other = new MailAddress("davische@teac.com.cn");
            //message.CC.Add(other); //添加抄送人
            //创建一个SmtpClient 类的新实例,并初始化实例的SMTP 事务的服务器
            SmtpClient client = new SmtpClient(@"192.168.16.254");
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.EnableSsl = false;
            //身份认证
            client.Credentials = new System.Net.NetworkCredential(mailfromname, mailpwd);
            bool ret = true; //返回值
            try
            {
                client.Send(message);
            }
            catch (SmtpException ex)
            {
                MessageBox.Show(ex.Message);
                ret = false;
            }
            catch (Exception ex2)
            {
                MessageBox.Show(ex2.Message);
                ret = false;
            }
            return ret;
        }
        #region
        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="mailTo">要发送的邮箱</param>
        /// <param name="mailSubject">邮箱主题</param>
        /// <param name="mailContent">邮箱内容</param>
        /// <returns>返回发送邮箱的结果</returns>
        public static bool SendEmail(string mailTo, string mailSubject, string mailContent)
        {

            string APP_Path = Application.StartupPath;//获取启动了应用程序的可执行文件的路径，不包括可执行文件的名称            
            string XmlFile = APP_Path + "\\ChangeMailto.xml";
            string mailto = Helper_Xml.Read(XmlFile, "/Root/SendtoList/From", "Mail").ToString();
            string mailtoname = Helper_Xml.Read(XmlFile, "/Root/SendtoList/From", "Name").ToString();
            string mailpwd = Helper_Xml.Read(XmlFile, "/Root/SendtoList/From", "Pwd").ToString();

            // 设置发送方的邮件信息,例如使用网易的smtp
            string smtpServer = "mail.teac.com.cn"; //SMTP服务器
            string mailFrom =  "\""+mailtoname+ "\" <" + mailto+">";//"\"DTA技术文管中心\" <ecnote@teac.com.cn> "; //登陆用户名
            string userPassword = mailpwd;//登陆密码

            // 邮件服务设置
            SmtpClient smtpClient = new SmtpClient();
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;//指定电子邮件发送方式
            smtpClient.Host = smtpServer; //指定SMTP服务器
            smtpClient.Credentials = new System.Net.NetworkCredential(mailFrom, userPassword);//用户名和密码

            // 发送邮件设置       
            MailMessage mailMessage = new MailMessage(mailFrom, mailTo); // 发送人和收件人
            mailMessage.Subject = mailSubject;//主题
            mailMessage.Body = mailContent;//内容
            mailMessage.BodyEncoding = Encoding.UTF8;// Encoding.GetEncoding(936);// Encoding.UTF8;//正文编码
            mailMessage.IsBodyHtml = false;//设置为HTML格式
            mailMessage.Priority = MailPriority.Low;//优先级

            try
            {
                smtpClient.Send(mailMessage); // 发送邮件
                return true;
            }
            catch (SmtpException)
            {
                return false;
            }
        }

        #endregion
        /// <summary> 
        /// 给定文件的路径，读取文件的二进制数据，判断文件的编码类型 
        /// </summary> 
        /// <param name=“FILE_NAME“>文件路径</param> 
        /// <returns>文件的编码类型</returns> 
        public static System.Text.Encoding GetType(string FILE_NAME)
        {
            FileStream fs = new FileStream(FILE_NAME, FileMode.Open, FileAccess.Read);
            Encoding r = GetType(fs);
            fs.Close();
            return r;
        }

        /// <summary> 
        /// 通过给定的文件流，判断文件的编码类型 
        /// </summary> 
        /// <param name=“fs“>文件流</param> 
        /// <returns>文件的编码类型</returns> 
        public static System.Text.Encoding GetType(FileStream fs)
        {
            byte[] Unicode = new byte[] { 0xFF, 0xFE, 0x41 };
            byte[] UnicodeBIG = new byte[] { 0xFE, 0xFF, 0x00 };
            byte[] UTF8 = new byte[] { 0xEF, 0xBB, 0xBF }; //带BOM 
            Encoding reVal = Encoding.Default;

            BinaryReader r = new BinaryReader(fs, System.Text.Encoding.Default);
            int i;
            int.TryParse(fs.Length.ToString(), out i);
            byte[] ss = r.ReadBytes(i);
            if (IsUTF8Bytes(ss) || (ss[0] == 0xEF && ss[1] == 0xBB && ss[2] == 0xBF))
            {
                reVal = Encoding.UTF8;
            }
            else if (ss[0] == 0xFE && ss[1] == 0xFF && ss[2] == 0x00)
            {
                reVal = Encoding.BigEndianUnicode;
            }
            else if (ss[0] == 0xFF && ss[1] == 0xFE && ss[2] == 0x41)
            {
                reVal = Encoding.Unicode;
            }
            r.Close();
            return reVal;

        }

        /// <summary> 
        /// 判断是否是不带 BOM 的 UTF8 格式 
        /// </summary> 
        /// <param name=“data“></param> 
        /// <returns></returns> 
        private static bool IsUTF8Bytes(byte[] data)
        {
            int charByteCounter = 1; //计算当前正分析的字符应还有的字节数 
            byte curByte; //当前分析的字节. 
            for (int i = 0; i < data.Length; i++)
            {
                curByte = data[i];
                if (charByteCounter == 1)
                {
                    if (curByte >= 0x80)
                    {
                        //判断当前 
                        while (((curByte <<= 1) & 0x80) != 0)
                        {
                            charByteCounter++;
                        }
                        //标记位首位若为非0 则至少以2个1开始 如:110XXXXX...........1111110X 
                        if (charByteCounter == 1 || charByteCounter > 6)
                        {
                            return false;
                        }
                    }
                }
                else
                {
                    //若是UTF-8 此时第一位必须为1 
                    if ((curByte & 0xC0) != 0x80)
                    {
                        return false;
                    }
                    charByteCounter--;
                }
            }
            if (charByteCounter > 1)
            {
                throw new Exception("非预期的byte格式");
            }
            return true;
        }



        /// <summary>
        /// GB2312转换成UTF8
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static string gb2312_utf8(string text)
        {
            //声明字符集   
            System.Text.Encoding utf8, gb2312;
            //gb2312   
            gb2312 = System.Text.Encoding.GetEncoding("gb2312");
            //utf8   
            utf8 = System.Text.Encoding.GetEncoding("utf-8");
            byte[] gb;
            gb = gb2312.GetBytes(text);
            gb = System.Text.Encoding.Convert(gb2312, utf8, gb);
            //返回转换后的字符   
            return utf8.GetString(gb);
        }

        /// <summary>
        /// UTF8转换成GB2312
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static string utf8_gb2312(string text)
        {
            //声明字符集   
            System.Text.Encoding utf8, gb2312;
            //utf8   
            utf8 = System.Text.Encoding.GetEncoding("utf-8");
            //gb2312   
            gb2312 = System.Text.Encoding.GetEncoding("gb2312");
            byte[] utf;
            utf = utf8.GetBytes(text);
            utf = System.Text.Encoding.Convert(utf8, gb2312, utf);
            //返回转换后的字符   
            return gb2312.GetString(utf);
        }
        /// <summary>
        /// 取得一个文本文件的编码方式。如果无法在文件头部找到有效的前导符，Encoding.Default将被返回。
        /// </summary>
        /// <param name="fileName">文件名。</param>
        /// <returns></returns>
        public static Encoding GetEncoding(string fileName)
        {
            return GetEncoding(fileName, Encoding.Default);
        }

        /// <summary>
        /// 取得一个文本文件流的编码方式。
        /// </summary>
        /// <param name="stream">文本文件流。</param>
        /// <returns></returns>
        public static Encoding GetEncoding(FileStream stream)
        {
            return GetEncoding(stream, Encoding.Default);
        }

        /// <summary>
        /// 取得一个文本文件的编码方式。
        /// </summary>
        /// <param name="fileName">文件名。</param>
        /// <param name="defaultEncoding">默认编码方式。当该方法无法从文件的头部取得有效的前导符时，将返回该编码方式。</param>
        /// <returns></returns>
        public static Encoding GetEncoding(string fileName, Encoding defaultEncoding)
        {
            FileStream fs = new FileStream(fileName, FileMode.Open);
            Encoding targetEncoding = GetEncoding(fs, defaultEncoding);
            fs.Close();
            return targetEncoding;
        }

        /// <summary>
        /// 取得一个文本文件流的编码方式。
        /// </summary>
        /// <param name="stream">文本文件流。</param>
        /// <param name="defaultEncoding">默认编码方式。当该方法无法从文件的头部取得有效的前导符时，将返回该编码方式。</param>
        /// <returns></returns>
        public static Encoding GetEncoding(FileStream stream, Encoding defaultEncoding)
        {
            Encoding targetEncoding = defaultEncoding;
            if (stream != null && stream.Length >= 2)
            {
                //保存文件流的前4个字节
                byte byte1 = 0;
                byte byte2 = 0;
                byte byte3 = 0;
                byte byte4 = 0;
                //保存当前Seek位置
                long origPos = stream.Seek(0, SeekOrigin.Begin);
                stream.Seek(0, SeekOrigin.Begin);

                int nByte = stream.ReadByte();
                byte1 = Convert.ToByte(nByte);
                byte2 = Convert.ToByte(stream.ReadByte());
                if (stream.Length >= 3)
                {
                    byte3 = Convert.ToByte(stream.ReadByte());
                }
                if (stream.Length >= 4)
                {
                    byte4 = Convert.ToByte(stream.ReadByte());
                }

                //根据文件流的前4个字节判断Encoding
                //Unicode {0xFF, 0xFE};
                //BE-Unicode {0xFE, 0xFF};
                //UTF8 = {0xEF, 0xBB, 0xBF};
                if (byte1 == 0xFE && byte2 == 0xFF)//UnicodeBe
                {
                    targetEncoding = Encoding.BigEndianUnicode;
                }
                if (byte1 == 0xFF && byte2 == 0xFE && byte3 != 0xFF)//Unicode
                {
                    targetEncoding = Encoding.Unicode;
                }
                if (byte1 == 0xEF && byte2 == 0xBB && byte3 == 0xBF)//UTF8
                {
                    targetEncoding = Encoding.UTF8;
                }

                //恢复Seek位置      
                stream.Seek(origPos, SeekOrigin.Begin);
            }
            return targetEncoding;
        }
    }
}
