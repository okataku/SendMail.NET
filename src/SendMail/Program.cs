using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;

namespace SendMail
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                WriteUsage();
                return;
            }


            var log = "";
            try
            {
                Console.WriteLine("Loading message informations....");

                var msg = Message.LoadFrom(args);
                var vars = new string[0];

                // [/v]オプションの値をパースする。
                {
                    var i = (args.ToList()).IndexOf("/v");
                    if (i > -1 && i < args.Length - 1)
                    {
                        vars = parseArgs(args[i + 1]);
                    }
                }

                // [/l]オプション
                {
                    var i = (args.ToList()).IndexOf("/l");
                    if (i > -1 && i < args.Length - 1)
                    {
                        log = args[i + 1];
                    }
                }


                //
                // メール送信処理
                //
                Console.WriteLine("Sending email....");
                var sc = new SmtpClient();
                sc.Host = msg.Host;
                sc.Port = msg.Port;
                sc.Timeout = msg.Timeout;
                sc.EnableSsl = msg.Ssl;
                sc.Credentials = new NetworkCredential(msg.User, msg.Password);

                var mm = new MailMessage();
                mm.From = msg.From;
                mm.SubjectEncoding = msg.SubjectEncoding;
                mm.Subject = vars.Length == 0 ? msg.Subject : String.Format(msg.Subject, vars);
                mm.IsBodyHtml = msg.IsBodyHtml;
                mm.BodyEncoding = msg.BodyEncoding;
                mm.Body = vars.Length == 0 ? msg.Body : String.Format(msg.Body, vars);
                foreach (var addr in msg.To) mm.To.Add(addr);
                foreach (var addr in msg.Cc) mm.CC.Add(addr);
                foreach (var addr in msg.Bcc) mm.Bcc.Add(addr);
                foreach (var atch in msg.Attachments) mm.Attachments.Add(atch);

                sc.Send(mm);
                Console.WriteLine("Finished sending.");
            }
            catch (Exception ex)
            {
                Func<Exception, string> w = null;
                w = new Func<Exception, string>(x =>
                {
                    return String.Format(">>Error {0:yyyy-MM-dd HH:mm:ss}\r\n[Type]{1}\r\n[Message]{2}\r\n[StackTrace]{3}\r\n[args]{4}",
                        DateTime.Now, x.GetType().FullName, x.Message, x.StackTrace, String.Join(" ", args)) +
                        (x.InnerException == null ? "" : w(x.InnerException));
                });

                var msg = w(ex);
                try
                {
                    // ログファイルに追記。
                    if (!String.IsNullOrEmpty(log))
                    {
                        using (var sw = new StreamWriter(log, true))
                        {
                            sw.WriteLine(msg);
                        }
                    }
                }
                catch(Exception ex1)
                {
                    var m = w(ex1);
                    Console.WriteLine(m);
                    Console.Error.WriteLine(m);
                    Environment.ExitCode = 1;
                }

                Console.WriteLine(msg);
                Console.Error.WriteLine(msg);
                Environment.ExitCode = 1;
            }
        }

        /// <summary>
        /// ';'区切りの値セットをパースします。
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        static string[] parseArgs(string args)
        {
            List<string> list = new List<string>();
            string token = "";
            for (int i = 0; i < args.Length; i++)
            {
                var chr = args[i];
                var lix = args.Length - 1;
                
                if (chr == ';')
                {
                    if (i != lix) {
                        list.Add(token);
                        token = "";
                    }
                }
                else if (chr == '\\')
                {
                    if (i != lix && args[i + 1] == ';')
                    {
                        token += ';';
                        i++;
                    }
                    else
                    {
                        token += '\\';
                    }
                }
                else
                {
                    token += chr;
                }

                if (i == lix)
                {
                    list.Add(token);
                }
            }

            return list.ToArray();
        }


        /// <summary>
        /// 使用方法を標準出力します。
        /// </summary>
        static void WriteUsage()
        {
            Console.WriteLine("SendMail");
            Console.WriteLine("");
            Console.WriteLine("  [Option]");
            Console.WriteLine("  /H hostname         Smtp server's host name or IP address.");
            Console.WriteLine("  /P 465              Smtp server's port.");
            Console.WriteLine("  /S                  Specify whether using ssl or not.");
            Console.WriteLine("  /U user:password    Username for authentication.");
            Console.WriteLine("  /T 6000             Timeout(ms).");
            Console.WriteLine("  /e utf-8            Encoding of subject.");
            Console.WriteLine("  /s subject          Subject.");
            Console.WriteLine("  /f foo@var.com<foo> From address.");
            Console.WriteLine("  /t foo@var.com<foo> To addresses. If you want to specify two or more");
            Console.WriteLine("                      addresses, join them by semicolon.");
            Console.WriteLine("  /c foo@var.com<foo> Cc addresses. If you want to specify two or more");
            Console.WriteLine("                      addresses, join them by semicolon.");
            Console.WriteLine("  /b foo@var.com<foo> Bcc addresses. If you want to specify two or more");
            Console.WriteLine("                      addresses, join them by semicolon.");
            Console.WriteLine("  /h                  Specify whether body is html or not.");
            Console.WriteLine("  /E utf-8            Encoding of body.");
            Console.WriteLine("  /B body             Body.");
            Console.WriteLine("  /a filepath         Attachments. If you want to specify two or more");
            Console.WriteLine("                      files, join them by semicolon.");
            Console.WriteLine("  /x filepath         MailMessage.xml.");
            Console.WriteLine("  /v arg1;arg2;...    Arguments for subject and body.");
            Console.WriteLine("  /l path/to/log      File path to error log file.");
        }
    }
}
