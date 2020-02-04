using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading;

namespace BackupReporter
{
    internal class Program
    {
        private static string targetArchiveName = string.Empty;

        private static void Main(string[] args)
        {
            try
            {
                Log("Uygulama başlatıldı");
                var appSettings = ConfigurationManager.AppSettings;
                string path = appSettings["BackupDirectory"];
                if (Directory.Exists(path))
                {
                    Log(string.Format("{0} dizinindeki son dosya bilgileri alınıyor.", path));
                    FileInfo newestFile = GetNewestFile(new DirectoryInfo(@path));
                    if (newestFile != null)
                    {
                        DateTime createDate = newestFile.CreationTime;
                        DateTime nowDate = DateTime.Now;
                        Log(string.Format("{0} dosyası seçildi. Dosya oluşturulma tarihi: {1}", newestFile.Name, createDate.ToLongTimeString()));
                        if (createDate.Year == nowDate.Year && createDate.Month == nowDate.Month && createDate.Day == nowDate.Day)
                        {
                            Rar(path, newestFile);
                            MailGonder(targetArchiveName, true);
                            DeleteArsive();
                        }
                        else
                        {
                            MailGonder(string.Empty, false);
                        }
                    }
                    else
                    {
                        Log("Dosya bulunamadı");
                    }
                }
                else
                {
                    Log("Dizin bulunamadı");
                }
            }
            catch (Exception ex)
            {
                Log(ex.Message);
            }
            finally
            {
                Log("Uygulama bitti");
            }
        }

        private static void DeleteArsive()
        {
            try
            {
                var appSettings = ConfigurationManager.AppSettings;
                string isDeleteArsive = "0";
                int configValue = 5;
                if (appSettings.AllKeys.Contains("IsDeleteArsive"))
                {
                    isDeleteArsive = appSettings["IsDeleteArsive"];
                }

                if (appSettings.AllKeys.Contains("SleepTime"))
                {
                    configValue = Convert.ToInt32(appSettings["SleepTime"]);
                }

                int sleepTime = configValue;

                if (isDeleteArsive == "1")
                {
                    Log("Arşiv temizleme işlemi yapılıyor...");
                    Thread.Sleep(new TimeSpan(0, sleepTime, 0));
                    if (!string.IsNullOrEmpty(targetArchiveName))
                    {
                        if (File.Exists(targetArchiveName))
                        {
                            File.Delete(targetArchiveName);
                        }
                    }
                    Log("Arşiv temizleme işlemi tamamlandı.");
                }
            }
            catch (Exception ex)
            {
                Log("DeleteArsive Metodunda hata oluştu." + ex.Message + ex.StackTrace);
            }
        }

        private static void MailGonder(string file, bool isSuccess)
        {
            try
            {
                Log("Mail gönderim işlemine geçildi.");

                var appSettings = ConfigurationManager.AppSettings;
                string smtpServer = appSettings["SmtpServer"];
                int smtpPort = Convert.ToInt32(appSettings["SmtpPort"]);
                string FromMail = appSettings["FromMail"];
                string FromMailDisplayName = appSettings["FromMailDisplayName"];
                string FromMailPassword = appSettings["FromMailPassword"];
                string ToMail = appSettings["ToMail"];
                string Subject = appSettings["Subject"];
                string Body = appSettings["Body"];

                SmtpClient sc = new SmtpClient();
                sc.Port = smtpPort;
                sc.Host = smtpServer;
                sc.EnableSsl = true;
                sc.Credentials = new NetworkCredential(FromMail, FromMailPassword);
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(FromMail, FromMailDisplayName);
                mail.To.Add(ToMail);
                mail.Subject = Subject;
                mail.IsBodyHtml = true;

                if (isSuccess)
                {
                    if (File.Exists(targetArchiveName))
                    {
                        mail.Attachments.Add(new Attachment(file));
                    }
                    else
                    {
                        Log("Maile eklenecek arşiv dosyası bulunamadı");
                    }
                    mail.Body = Body;
                }
                else
                {
                    mail.Body = appSettings["NoBackupMessage"];
                }
                sc.Send(mail);
            }
            catch (Exception ex)
            {
                Log(string.Format("Mail gönderim işleminde hata oluştu. Detay:{0} {1}", ex.Message, ex.StackTrace));
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Log("Mail gönderim işlemleri tamamlandı.");
            }
        }

        private static void Rar(string directory, FileInfo file)
        {
            Log("Rar işlemine geçildi.");

            try
            {
                string targetFile = directory + file.Name;
                string fileName = file.Name.Split('.')[0];
                targetArchiveName = directory + fileName + ".rar";
                string currentDirectory = Environment.CurrentDirectory;
                ProcessStartInfo startInfo = new ProcessStartInfo(currentDirectory + "\\Rar.exe");
                startInfo.CreateNoWindow = true;
                startInfo.RedirectStandardOutput = true;
                startInfo.UseShellExecute = false;
                startInfo.WindowStyle = ProcessWindowStyle.Maximized;
                startInfo.Arguments = string.Format(" a -r -ep \"{0}\" \"{1}\"", targetArchiveName, targetFile);

                using (Process exeProcess = Process.Start(startInfo))
                {
                    exeProcess.WaitForExit();
                    exeProcess.Close();
                }
            }
            catch (Exception ex)
            {
                Log(string.Format("Rar işleminde hata oluştu. Detay:{0} {1}", ex.Message, ex.StackTrace));
            }
            finally
            {
                Log("Rar işlemleri tamamlandı.");
            }
        }

        public static FileInfo GetNewestFile(DirectoryInfo directory)
        {
            if (directory.GetFiles().Length > 0)
            {
                return directory.GetFiles()
                    .Union(directory.GetDirectories().Select(d => GetNewestFile(d)))
                    .OrderByDescending(f => (f == null ? DateTime.MinValue : f.LastWriteTime))
                    .FirstOrDefault();
            }

            return null;
        }

        private static void Log(string message)
        {
            try
            {
                string fileName = DateTime.Now.ToString("yyyyddMM") + ".txt";
                using (StreamWriter w = File.AppendText(fileName))
                {
                    w.WriteLine(string.Format("{0}|{1}", DateTime.Now.ToString("yyyyMMddHHmmss"), message));
                    Console.WriteLine(message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}