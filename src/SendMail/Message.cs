using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Net.Mail;
using System.Xml;

namespace SendMail
{
    class Message
    {
        /// <summary>
        /// SMTPサーバーのホスト名を取得または設定します。
        /// </summary>
        public string Host { get; set; }
        /// <summary>
        /// SMTPサーバーのポートを取得または設定します。
        /// </summary>
        public int Port { get; set; }
        /// <summary>
        /// SSL通信を利用するかを取得または設定します。
        /// </summary>
        public bool Ssl { get; set; }
        /// <summary>
        /// メールアカウントのユーザー名を取得または設定します。
        /// </summary>
        public string User { get; set; }
        /// <summary>
        /// メールアカウントのパスワードを取得または設定します。
        /// </summary>
        public string Password { get; set; }
        /// <summary>
        /// メール送信処理のタイムアウト時間(ms)を取得または設定します。
        /// </summary>
        public int Timeout { get; set; }

        /// <summary>
        /// メールの送信元アドレスを取得または設定します。
        /// </summary>
        public MailAddress From { get; set; }
        /// <summary>
        /// メールの送信先を取得または設定します。
        /// </summary>
        public ICollection<MailAddress> To { get; set; }
        /// <summary>
        /// メールの送信先を取得または設定します。
        /// </summary>
        public ICollection<MailAddress> Cc { get; set; }
        /// <summary>
        /// メールの送信先を取得または設定します。
        /// </summary>
        public ICollection<MailAddress> Bcc { get; set; }

        /// <summary>
        /// 件名の文字エンコーディングを取得または設定します。
        /// </summary>
        public Encoding SubjectEncoding { get; set; }
        /// <summary>
        /// 件名を取得または設定します。
        /// </summary>
        public string Subject { get; set; }
        /// <summary>
        /// 本文の文字エンコーディングを取得または設定します。
        /// </summary>
        public Encoding BodyEncoding { get; set; }
        /// <summary>
        /// 本文がHTML文書かどうかを取得または設定します。
        /// </summary>
        public bool IsBodyHtml { get; set; }
        /// <summary>
        /// 本文を取得または設定します。
        /// </summary>
        public string Body { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public ICollection<Attachment> Attachments { get; set; }

        /// <summary>
        /// NTTool.SendMail.Messageクラスの新しいインスタンスを初期化します。
        /// </summary>
        public Message()
        {
            this.Port = 25;
            this.To = new List<MailAddress>();
            this.Cc = new List<MailAddress>();
            this.Bcc = new List<MailAddress>();
            SubjectEncoding = Encoding.GetEncoding("utf-8");
            BodyEncoding = Encoding.GetEncoding("utf-8");
            this.Attachments = new List<Attachment>();
        }

        /// <summary>
        /// メッセージ定義XMLファイルから定義情報をロードします。
        /// </summary>
        /// <param name="file">メッセージ定義XMLファイルのパス。</param>
        /// <returns>NTTool.SendMail.Messageオブジェクト。</returns>
        public static Message LoadFrom(string file)
        {
            var doc = new XmlDocument();
            doc.Load(file);

            var msg = new Message();

            var node_smtp = doc.SelectSingleNode("sendmail/smtp");
            var node_msg = doc.SelectSingleNode("sendmail/message");

            // Host
            {
                var node = node_smtp.SelectSingleNode("host");
                if (node != null) msg.Host = node.InnerText;
            }

            // Port
            {
                var node = node_smtp.SelectSingleNode("port");
                if (node != null) msg.Port = int.Parse(node.InnerText);
            }

            // SSL
            {
                var node = node_smtp.SelectSingleNode("ssl");
                if (node != null) msg.Ssl = bool.Parse(node.InnerText);
            }

            // User
            {
                var node = node_smtp.SelectSingleNode("user");
                if (node != null) msg.User = node.InnerText;
            }

            // Password
            {
                var node = node_smtp.SelectSingleNode("password");
                if (node != null) msg.Password = node.InnerText;
            }

            // Timeout
            {
                var node = node_smtp.SelectSingleNode("timeout");
                if (node != null) msg.Timeout = int.Parse(node.InnerText);
            }

            // Subject
            {
                var node = node_msg.SelectSingleNode("subject");
                if (node != null)
                {
                    foreach (XmlAttribute attr in node.Attributes)
                    {
                        if (attr.Name == "encoding")
                        {
                            msg.Subject = attr.Value;
                        }
                    }
                    msg.Subject = node.InnerText;
                }
            }

            // From
            {
                var node = node_msg.SelectSingleNode("from/address");
                if (node != null)
                {
                    var name = "";
                    foreach (XmlAttribute attr in node.Attributes)
                    {
                        if (attr.Name == "displayName")
                        {
                            name = attr.Value;
                        }
                    }
                    msg.From = String.IsNullOrEmpty(name) ? new MailAddress(node.InnerText) : new MailAddress(node.InnerText, name);
                }
            }

            // To
            {
                var nodes = node_msg.SelectNodes("to/address");
                if (nodes != null)
                {
                    foreach (XmlNode node in nodes)
                    {
                        var name = "";
                        foreach (XmlAttribute attr in node.Attributes)
                        {
                            if (attr.Name == "displayName")
                            {
                                name = attr.Value;
                            }
                        }
                        msg.To.Add(String.IsNullOrEmpty(name) ? new MailAddress(node.InnerText) : new MailAddress(node.InnerText, name));
                    }
                }
            }

            // Cc
            {
                var nodes = node_msg.SelectNodes("cc/address");
                if (nodes != null)
                {
                    foreach (XmlNode node in nodes)
                    {
                        var name = "";
                        foreach (XmlAttribute attr in node.Attributes)
                        {
                            if (attr.Name == "displayName")
                            {
                                name = attr.Value;
                            }
                        }
                        msg.Cc.Add(String.IsNullOrEmpty(name) ? new MailAddress(node.InnerText) : new MailAddress(node.InnerText, name));
                    }
                }
            }

            // Bcc
            {
                var nodes = node_msg.SelectNodes("bcc/address");
                if (nodes != null)
                {
                    foreach (XmlNode node in nodes)
                    {
                        var name = "";
                        foreach (XmlAttribute attr in node.Attributes)
                        {
                            if (attr.Name == "displayName")
                            {
                                name = attr.Value;
                            }
                        }
                        msg.Bcc.Add(String.IsNullOrEmpty(name) ? new MailAddress(node.InnerText) : new MailAddress(node.InnerText, name));
                    }
                }
            }

            // Body
            {
                var node = node_msg.SelectSingleNode("body");
                if (node != null)
                {
                    foreach (XmlAttribute attr in node.Attributes)
                    {
                        if (attr.Name == "encoding")
                        {
                            msg.BodyEncoding = Encoding.GetEncoding(attr.Value);
                        }
                        else if (attr.Name == "html")
                        {
                            msg.IsBodyHtml = bool.Parse(attr.Value);
                        }
                    }
                    if (node.HasChildNodes && node.FirstChild.NodeType == XmlNodeType.CDATA)
                    {
                        msg.Body = node.FirstChild.Value;
                    }
                    else
                    {
                        msg.Body = node.InnerText;
                    }
                }
            }

            // Attachments
            {
                var nodes = node_msg.SelectNodes("attachment/file");
                if (nodes != null)
                {
                    foreach (XmlNode node in nodes)
                    {
                        msg.Attachments.Add(new Attachment(node.InnerText));
                    }
                }
            }

            return msg;
        }
        /// <summary>
        /// メッセージ定義XMLファイルから定義情報をロードします。
        /// </summary>
        /// <param name="args">実行時の引数。</param>
        /// <returns>NTTool.SendMail.Messageオブジェクト。</returns>
        public static Message LoadFrom(string[] args)
        {
            var a = args.ToList();
            var l = args.Length - 1;
            var msg = new Message();

            /* 
             * [x]オプションが指定された場合
             * メッセージ定義ファイルで定義されている
             * 内容を引数の値で上書きする。
             * 
             */


            // XML
            {
                var i = a.IndexOf("/x");
                if (i > -1 && i < l)
                {
                    msg = LoadFrom(a[i + 1]);
                }
            }

            // Host
            {
                var i = a.IndexOf("/H");
                if (i > -1 && i < l)
                {
                    msg.Host = a[i + 1];
                }
            }

            // Port
            {
                var i = a.IndexOf("/P");
                var p = 0;
                if (i > -1 && i < l && int.TryParse(a[i + 1], out p))
                {
                    msg.Port = p;
                }
            }

            // SSL
            {
                if (a.IndexOf("/S") > -1)
                {
                    msg.Ssl = true;
                }
            }

            // User:Password
            {
                var i = a.IndexOf("/U");
                if (i > 1 || i < l)
                {
                    var tokens = a[i + 1].Split(new char[] { ':' });
                    if (tokens.Length > 0) msg.User = tokens[0];
                    if (tokens.Length > 1) msg.Password = tokens[1];
                }
            }

            // Timeout
            {
                var i = a.IndexOf("/T");
                var t = 0;
                if (i > -1 && i < l && int.TryParse(a[i + 1], out t))
                {
                    msg.Timeout = t;
                }
            }

            // Subject Encoding
            {
                var i = a.IndexOf("/e");
                if (i > -1 && i < l)
                {
                    msg.SubjectEncoding = Encoding.GetEncoding(a[i + 1]);
                }
            }

            // Subject
            {
                var i = a.IndexOf("/s");
                if (i > -1 && i < l)
                {
                    msg.Subject = a[i + 1];
                }
            }

            // From
            {
                var i = a.IndexOf("/f");
                if (i > -1 && i < l)
                {
                    var rgx = new Regex(@"^?(.+?)\<(.+?)\>$");
                    var ad = a[i + 1];
                    var match = rgx.Match(ad);
                    msg.From = match.Success ?
                        new MailAddress(match.Groups[1].Value, match.Groups[2].Value) : new MailAddress(ad);
                }
            }

            // To
            {
                var i = a.IndexOf("/t");
                if (i > -1 && i < l)
                {
                    msg.To.Clear();
                    var rgx = new Regex(@"^?(.+?)\<(.+?)\>$");
                    foreach (var ad in a[i + 1].Split(new char[] { ';' }))
                    {
                        var match = rgx.Match(ad);
                        msg.To.Add(match.Success ?
                            new MailAddress(match.Groups[1].Value, match.Groups[2].Value) : new MailAddress(ad));
                    }

                }
            }

            // Cc
            {
                var i = a.IndexOf("/c");
                if (i > -1 && i < l)
                {
                    msg.Cc.Clear();
                    var rgx = new Regex(@"^?(.+?)\<(.+?)\>$");
                    foreach (var ad in a[i + 1].Split(new char[] { ';' }))
                    {
                        var match = rgx.Match(ad);
                        msg.Cc.Add(match.Success ?
                            new MailAddress(match.Groups[1].Value, match.Groups[2].Value) : new MailAddress(ad));
                    }

                }
            }

            // Bcc
            {
                var i = a.IndexOf("/b");
                if (i > -1 && i < l)
                {
                    msg.Bcc.Clear();
                    var rgx = new Regex(@"^?(.+?)\<(.+?)\>$");
                    foreach (var ad in a[i + 1].Split(new char[] { ';' }))
                    {
                        var match = rgx.Match(ad);
                        msg.Bcc.Add(match.Success ?
                            new MailAddress(match.Groups[1].Value, match.Groups[2].Value) : new MailAddress(ad));
                    }

                }
            }

            // HTML
            {
                if (a.IndexOf("/h") > -1)
                {
                    msg.IsBodyHtml = true;
                }
            }

            // Body Encoding
            {
                var i = a.IndexOf("/E");
                if (i > -1 && i < l)
                {
                    msg.BodyEncoding = Encoding.GetEncoding(a[i + 1]);
                }
            }

            // Body
            {
                var i = a.IndexOf("/B");
                if (i > -1 && i < l)
                {
                    msg.Body = a[i + 1];
                }
            }

            // Attachments
            {
                var i = a.IndexOf("/a");
                if (i > -1 && i < l)
                {
                    msg.Attachments.Clear();
                    var files = a[i + 1].Split(new char[] { ';' });
                    foreach (var file in files)
                    {
                        msg.Attachments.Add(new Attachment(file));
                    }
                }
            }

            return msg;
        }
    }

    class MessageLoadException : Exception
    {
        /// <summary>
        /// NTTool.SendMail.Messageクラスの新しいインスタンスを初期化します。
        /// </summary>
        public MessageLoadException()
            : base()
        {

        }
        /// <summary>
        /// NTTool.SendMail.Messageクラスの新しいインスタンスを初期化します。
        /// </summary>
        /// <param name="message"></param>
        public MessageLoadException(string message)
            : base(message)
        {

        }
        /// <summary>
        /// NTTool.SendMail.Messageクラスの新しいインスタンスを初期化します。
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public MessageLoadException(string message, Exception innerException)
            : base(message, innerException)
        {
            
        }
        /// <summary>
        /// NTTool.SendMail.Messageクラスの新しいインスタンスを初期化します。
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        public MessageLoadException(System.Runtime.Serialization.SerializationInfo info, System.Runtime.Serialization.StreamingContext context)
            : base(info, context)
        {

        }
    }
}
