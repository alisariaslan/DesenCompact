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
using System.Data.SqlTypes;
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
            read("notlarim.txt");
            if (!File.Exists("notlarim.txt"))
            {
                string text = "SPECS\n\n" + getspecs();
                write(text, "notlarim.txt");
                read("notlarim.txt");
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
            File.Copy("start.lnk", desktop_path + "\\DesenPOS.lnk");

            //string shortcut_path = (desktop_path + "\\DesenPOS.lnk");

            //IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
            //IWshRuntimeLibrary.IWshShortcut sc = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(shortcut_path);
            //sc.Description = "some desc";
            //sc.TargetPath = "C:\\DesenPOS\\DesenPos.exe";
            //sc.Save();
            log($"DesenPOS kısayolu düzeltildi.");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            log("UPDATE...");
            string sqlString = "USE DesenPOS;\nSELECT * FROM Sirket;";
            string Adi = "UPDATE Sirket SET Adi='" + textBox1.Text + "';";
            sqlString += "\n" + Adi;
            string VergiNo = "UPDATE Sirket SET VergiNo = '" + textBox2.Text + "';";
            sqlString += "\n" + VergiNo;
            string CustomerId = "UPDATE Sirket SET CustomerId = '" + textBox3.Text + "';";
            if (!textBox3.Text.Equals("")) {
                sqlString += "\n" + CustomerId;
                sqlString += "\n" + "SELECT Adi, VergiNo, CustomerId FROM Sirket;";
            } else
            {
                sqlString += "\n" + "SELECT Adi, VergiNo FROM Sirket;";
            }
            if (File.Exists("update.sql"))
                File.Delete("update.sql");
            write(sqlString, "update.sql");
            Process.Start("update.sql");
        }

        private async void button5_Click(object sender, EventArgs e)
        {
            Directory.CreateDirectory("surum");
            log("Alt klasör oluşturuldu: surum");
            log("Sürüm arşivi çıkartılıyor... surum.rar");
            await extract("surum.rar", "C:\\DesenPOS", progressBar1);
            log($"Sürüm atma işlemi tamamlandı.");
            Process.Start("C:\\DesenPOS\\versiyon.txt");

        }

        private void button6_Click(object sender, EventArgs e)
        {
            log("Database.sql dosyası çalıştırılıyor...");
            Process.Start("Database.sql");
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            write(richTextBox1.Text, "notlarim.txt");
        }

        private void write(string text, string path)
        {
            StreamWriter streamWriter = new StreamWriter(path, append: false);
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

        private void read(string path)
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

            Process.Start("start.lnk");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe", "kurulum");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe","C:\\DesenPOS");
        }
    }
}
