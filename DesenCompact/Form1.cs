using Microsoft.SqlServer.Management.Smo;
using SharpCompress.Archives;
using SharpCompress.Archives.Rar;
using SharpCompress.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
        string desktop_path;

        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
            desktop_path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
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
            string shortcut_path = (desktop_path+"\\DesenPOS.lnk");
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
         


            string Adi = textBox1.Text;
            string VergiNo = textBox2.Text;
            string CustomerId = textBox3.Text;
            string connectionString = "";
            string sqlString = "";
            
            //SqlConnection sqlConnection = new SqlConnection("connectionString");
            //SqlCommand sqlCommand = new SqlCommand(sqlString,sqlConnection);
            //if (sqlConnection.State != ConnectionState.Open)
            //    sqlConnection.Open();
            //sqlCommand.ExecuteNonQuery();
            //sqlConnection.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DataTable dataTable = SmoApplication.EnumAvailableSqlServers(true);
            listBox2.ValueMember = "Name";
            listBox2.DataSource = dataTable;
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            if (listBox2.SelectedIndex != -1)
            {
                string serverName = listBox2.SelectedValue.ToString();
                Server server = new Server(serverName);
                try
                {
                    foreach (Database database in server.Databases)
                    {
                        listBox3.Items.Add(database.Name);
                    }
                }
                catch (Exception ex)
                {
                    string exception = ex.Message;
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
