using Hardware.Info;
using Microsoft.SqlServer.Management.Smo;
using Microsoft.VisualBasic.Devices;
using Microsoft.Win32;
using SharpCompress.Archives;
using SharpCompress.Archives.Rar;
using SharpCompress.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DesenCompact
{
    public partial class Form1 : Form
    {
        SqlConnection sqlConnection;
        string desktop_path;
        string sql_connection_string;

        public Form1()
        {
            //File.Delete("notlarim.txt");
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
            desktop_path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            sqlConnection = new SqlConnection("");
            sql_connection_string = @"Data Source=.\DESENERP; Initial Catalog=DesenPOS; uid=sa; password=DesenErp.12345;";
            read();
            if (!File.Exists("notlarim.txt"))
            {
                string text = "SPECS\n\n" + getspecs();
                write(text);
                read();
            }
        }

        private void log(string message)
        {
            String hourMinute = DateTime.Now.ToString("HH:mm:ss");
            listBox1.Items.Add($"{message} ({hourMinute})");
        }

        public static async Task extract(string filefullpath, string write_path, ProgressBar progressBar)
        {
            await Task.Run(async () =>
            {
                using (var archive = RarArchive.Open(filefullpath))
                {
                    progressBar.Maximum = archive.Entries.Where(entry => !entry.IsDirectory).Count();
                    progressBar.Value = 0;
                    progressBar.Minimum = 0;
                    foreach (var entry in archive.Entries.Where(entry => !entry.IsDirectory))
                    {
                        progressBar.Value += 1;
                        entry.WriteToDirectory(write_path + "/", new ExtractionOptions()
                        {
                            ExtractFullPath = true,
                            Overwrite = true
                        });
                    }
                }

            });
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string cplPath = Path.Combine(Environment.SystemDirectory, "control.exe");
            Process.Start(cplPath, "/name Microsoft.ProgramsAndFeatures");
            log("Windows özellikleri için denetim masası açıldı.");
        }
        private async void button2_Click(object sender, EventArgs e)
        {
            Directory.CreateDirectory("kurulum");
            log("Alt klasör oluşturuldu: kurulum");
            log("Kurulum arşivi çıkartılıyor... kurulum.rar");
            await extract("kurulum.rar", "kurulum", progressBar1);
            log($"Arşiv çıkarma işlemi tamamlandı.");
            Process.Start("explorer.exe", "kurulum");
            log("Kurulum klasörü açıldı.");
        }

        private async void button3_Click(object sender, EventArgs e)
        {
            log("Malzeme arşivi C:\\DesenPOS yoluna çıkartılıyor... malzeme.rar");
            await extract("malzeme.rar", "C:\\DesenPOS", progressBar1);
            log($"Arşiv çıkarma işlemi tamamlandı.");
            string shortcut_path = (desktop_path + "\\DesenPOS.lnk");
            IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
            IWshRuntimeLibrary.IWshShortcut sc = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(shortcut_path);
            sc.Description = "some desc";
            sc.TargetPath = "C:\\DesenPOS\\DesenPos.exe";
            sc.Save();
            log($"DesenPOS kısayolu düzeltildi.");
            Process.Start(desktop_path + "\\DesenPOS.lnk");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            log("UPDATE...");
            string Adi = textBox1.Text;
            string VergiNo = textBox2.Text;
            string CustomerId = textBox3.Text;
            string sqlString = "use DesenPOS;" +
                " DELETE FROM Sirket;" +
                " INSERT INTO Sirket(Adi) VALUES('" + Adi + "');" +
                " UPDATE Sirket SET VergiNo = '" + VergiNo + "';" +
                " UPDATE Sirket SET CustomerId = '" + CustomerId + "';" +
                " SELECT Adi, VergiNo, CustomerId FROM Sirket; ";
            SqlCommand sqlCommand = new SqlCommand(sqlString, sqlConnection);
            int numberof_rows_effected = sqlCommand.ExecuteNonQuery();
            log(numberof_rows_effected + " satır etkilendi.");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            string ServerName = Environment.MachineName;
            RegistryView registryView = Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32;
            using (RegistryKey hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, registryView))
            {
                RegistryKey instanceKey = hklm.OpenSubKey(@"SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL", false);
                if (instanceKey != null)
                {
                    foreach (var instanceName in instanceKey.GetValueNames())
                    {
                        listBox2.Items.Add(ServerName + "\\" + instanceName);
                    }
                }
            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            sql_connection_string = @"Data Source=" + listBox2.SelectedItem.ToString() + "; Initial Catalog=DesenPOS; uid=sa; password=DesenErp.12345;";
            textBox4.Text = sql_connection_string;
            sqlConnection = new SqlConnection(sql_connection_string);
            sqlConnection.Close();
            try
            {
                sqlConnection.Open();
            }
            catch (Exception ex)
            {
                log(ex.Message);
            }
            if (sqlConnection.State == ConnectionState.Open)
                log("Bağlantı başarılı.");
            else
                log("HATA! Bağlantı yapılamadı!");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Process.Start("database.sql");
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            write(richTextBox1.Text);
        }

        private void write(string text)
        {
            StreamWriter streamWriter = new StreamWriter("notlarim.txt", append: false);
            streamWriter.Write(text);
            streamWriter.Close();
        }

        private string getspecs()
        {
            ComputerInfo computerInfo = new ComputerInfo();
            IHardwareInfo hardwareInfo = new HardwareInfo();
            hardwareInfo.RefreshMotherboardList();
            hardwareInfo.RefreshCPUList();
            hardwareInfo.RefreshMemoryStatus();
            string text = "";
            text = "İşletim Sistemi: " + computerInfo.OSFullName;
            text += "\n\nAnakart: " + hardwareInfo.MotherboardList[0];
            text += "\nİşlemci: " + hardwareInfo.CpuList[0];
            text += "\nRAM: Toplam " + byteToGB((long)computerInfo.TotalPhysicalMemory);
            foreach (DriveInfo drive in DriveInfo.GetDrives())
            {
                if (drive.IsReady)
                {
                    text += "\nDisk " + drive.Name + ": Kullanılabilir alan " + byteToGB(drive.TotalFreeSpace);
                }
            }
            return text;
        }

        private string byteToGB(long bytes)
        {
            string[] Suffix = { "B", "KB", "MB", "GB", "TB" };
            int i;
            double dblSByte = bytes;
            for (i = 0; i < Suffix.Length && bytes >= 1024; i++, bytes /= 1024)
            {
                dblSByte = bytes / 1024.0;
            }
            return String.Format("{0:0.##} {1}", dblSByte, Suffix[i]);
        }

        private void read()
        {
            if (File.Exists("notlarim.txt"))
            {
                StreamReader streamReader = new StreamReader("notlarim.txt");
                string text = "";
                while (!streamReader.EndOfStream)
                {
                    text += streamReader.ReadLine() + "\n";
                }
                richTextBox1.Text = text;
                streamReader.Close();
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            sqlConnection.Close();
        }

        private void Form1_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            Process.Start("yardim.pdf");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Process.Start("C:\\DesenPOS\\DesenPos.exe");
        }
    }
}
