using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;
using System.Text.RegularExpressions;


namespace WTF1._0
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string currentdirectory = Directory.GetCurrentDirectory();
        private void Form1_Load(object sender, EventArgs e)
        {
            //var pc = new PrincipalContext(ContextType.Machine);
            //var upGUID = (UserPrincipal.FindByIdentity(pc, "Оператор")).Sid;
            //var gpGUID = (GroupPrincipal.FindByIdentity(pc, "Администраторы"));

            string r = ("AuditPol /backup /file:\"" + currentdirectory + "\\auditpolicy.csv\"");
            var procproc = System.Diagnostics.Process.Start("cmd.exe", "/C" + r);
            while (!procproc.HasExited) ;

            string g = ("secedit /export /cfg \"" + currentdirectory + "\\secpol.cfg\"");
            var proc = System.Diagnostics.Process.Start("cmd.exe", "/C" + g);
            while (!proc.HasExited) ;
            
            zapolninie(dataGridView1, Directory.GetCurrentDirectory() + "\\StandardАдминистраторы.csv", "Администраторы");
            zapolninie(dataGridView2, Directory.GetCurrentDirectory() + "\\StandardНеадминистраторы.csv","Не администраторы");
            zapolninie(dataGridView3, Directory.GetCurrentDirectory() + "\\StandardАдминистратор.csv", "Администратор");
            zapolninie(dataGridView4, Directory.GetCurrentDirectory() + "\\StandardКураторИБ.csv", "Куратор ИБ");
            zapolninie(dataGridView3, Directory.GetCurrentDirectory() + "\\StandardОператор.csv", "Оператор");

        } /*при заупске приложения, будут сразу исполняться его функции */
        public void zapolninie(DataGridView dataGridViewb, string path, string user)
        {

            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("Номер Политики");
            dataTable.Columns.Add("Путь");
            dataTable.Columns.Add("Эталонное Значение");
            dataTable.Columns.Add("Текущее Значение");
            dataTable.Columns.Add("Статус");

            if (File.Exists(path))
            {
                StreamReader streamReader = new StreamReader(path, Encoding.Default);
                string[] totalData = new string[File.ReadAllLines(path).Length];
                while (!streamReader.EndOfStream)
                {
                    totalData = streamReader.ReadLine().Split(';');
                    dataTable.Rows.Add(totalData[0], totalData[1], totalData[2]);
                }
                dataGridViewb.DataSource = dataTable;

                int i = dataGridViewb.RowCount;
                for (int j = 0; j < i; j++)
                {
                    poisk(dataGridViewb.Rows[j].Cells[1].Value.ToString(), j, dataGridViewb, user);
                }
                cvet(dataGridViewb);

            }
        }

        public void cvet(DataGridView dataGridViewb)
        {
            int uspex = 0;
            int noparametr = 0; ;
            int neuspex = 0;
            int i = dataGridViewb.RowCount;
            for (int j = 0; j < i; j++)
            {
                if (dataGridViewb.Rows[j].Cells[2].Value.ToString() == "")
                {
                    dataGridViewb.Rows[j].Cells[4].Value = "Параметр отсутствует в ОС";
                    noparametr++;
                }
                else
                {
                    if (dataGridViewb.Rows[j].Cells[2].Value.ToString() != dataGridViewb.Rows[j].Cells[3].Value.ToString())
                    {
                        dataGridViewb.Rows[j].Cells[4].Value = "Не успех";
                        dataGridViewb.Rows[j].DefaultCellStyle.BackColor = Color.PaleVioletRed;
                        neuspex++;
                    }

                    else if ((dataGridViewb.Rows[j].Cells[2].Value.ToString() == dataGridViewb.Rows[j].Cells[3].Value.ToString()))
                    {
                        dataGridViewb.Rows[j].Cells[4].Value = "Успех";
                        dataGridViewb.Rows[j].DefaultCellStyle.BackColor = Color.LightGreen;
                        uspex++;
                    }
                }
                label1.Text = "Успех - " + uspex;
                label2.Text = "Не успех - " + neuspex;
                label3.Text = "Параметр отсутствует - " + noparametr;
                label4.Text = "Имя пользователя: " + System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            }
        }

        public void poisk(string parametr, int Stroka, DataGridView dataGridViewb, string user)
        {
            if (Regex.IsMatch(parametr, "HKEY"))
                PoiskPoRegedit(parametr, Stroka, dataGridViewb,user);
            else if (Regex.IsMatch(parametr, "secedit"))
                PoiskPoCmdSecedit(parametr, Stroka, dataGridViewb);
            else if (Regex.IsMatch(parametr, "AuditPol"))
                PoiskPoCmdAuditPol(parametr, Stroka, dataGridViewb);
        } /*поиск выполняется в 3 директориях: реестр, политика аудитов, политика безопасности*/

        public void PoiskPoRegedit(string parametr, int Stroka, DataGridView dataGridView1, string User)
        {

            //var pc = new PrincipalContext(ContextType.Machine);
            ////var upGUID = (UserPrincipal.FindByIdentity(pc, User).GetUnderlyingObject() as DirectoryEntry).Guid;
            ////var gpGUID = (GroupPrincipal.FindByIdentity(pc, User).GetUnderlyingObject() as DirectoryEntry).Guid;
            //var upSID = (UserPrincipal.FindByIdentity(pc, "Оператор")).Sid;
            //var gpSID = (GroupPrincipal.FindByIdentity(pc, "Администраторы")).Sid;


            string[] words = parametr.Split(':');

            if (Regex.IsMatch(parametr, "HKEY_LOCAL_MACHINE"))
            {
                string pobeda = words[0].Replace("HKEY_LOCAL_MACHINE\\", "");
                if (Registry.LocalMachine.OpenSubKey(pobeda) != null)
                    dataGridView1.Rows[Stroka].Cells[3].Value = Registry.GetValue(words[0], words[1], "Параметр отсутствует").ToString();
                else
                    dataGridView1.Rows[Stroka].Cells[3].Value = "Отсутствует директория параметра";
            }
            else if (Regex.IsMatch(parametr, "HKEY_CURRENT_USER"))
            {
                string pobeda = words[0].Replace("HKEY_CURRENT_USER\\", "");
                if (Registry.CurrentUser.OpenSubKey(pobeda) != null)
                    dataGridView1.Rows[Stroka].Cells[3].Value = Registry.GetValue(@words[0], words[1], "Параметр отсутствует").ToString();
                else
                    dataGridView1.Rows[Stroka].Cells[3].Value = "Отсутствует директория параметра";
            }
            else if (Regex.IsMatch(parametr, "HKEY_USERS"))
            {
                string pobeda = words[0].Replace("HKEY_USERS\\", "");
                if (Registry.Users.OpenSubKey(pobeda).ToString() != "")
                    dataGridView1.Rows[Stroka].Cells[3].Value = Registry.GetValue(words[0], words[1], "Параметр отсутствует").ToString();
                else
                    dataGridView1.Rows[Stroka].Cells[3].Value = "Отсутствует директория параметра";
            }
            else if (Regex.IsMatch(parametr, "HKEY_CLASSES_ROOT"))
            {
                string pobeda = words[0].Replace("HKEY_CLASSES_ROOT\\", "");
                if (Registry.ClassesRoot.OpenSubKey(pobeda) != null)
                    dataGridView1.Rows[Stroka].Cells[3].Value = Registry.GetValue(words[0], words[1], "Параметр отсутствует").ToString();
                else
                    dataGridView1.Rows[Stroka].Cells[3].Value = "Отсутствует директория параметра";
            }
            else if (Regex.IsMatch(parametr, "HKEY_CURRENT_CONFIG"))
            {
                string pobeda = words[0].Replace("HKEY_CURRENT_CONFIG\\", "");
                if (Registry.CurrentConfig.OpenSubKey(pobeda) != null)
                    dataGridView1.Rows[Stroka].Cells[3].Value = Registry.GetValue(words[0], words[1], "Параметр отсутствует").ToString();
                else
                    dataGridView1.Rows[Stroka].Cells[3].Value = "Отсутствует директория параметра";
            }
        } /*поиск по всему реестру, эталонные значения задавать в десятичной системе*/

        public void PoiskPoCmdSecedit(string parametr, int Stroka, DataGridView dataGridView1)
        {
            StreamReader sr = new StreamReader(currentdirectory + "\\secpol.cfg", Encoding.Default);
            string pobeda = File.ReadAllText(currentdirectory + "\\secpol.cfg").Replace("\r", "");
            string tidumaletovso = pobeda.Replace(" ", "");
            string[] pobeda2 = tidumaletovso.Split(']');
            string[] words = (pobeda2[2] + pobeda2[3]).Split('\n');

            string lol = parametr.Replace(" ", "");
            string[] kek = lol.Split('|');

            foreach (string s in words)
            {
                string[] skolkouzhemozhno = s.Split('=');
                if (Regex.IsMatch(s, kek[1]))
                    dataGridView1.Rows[Stroka].Cells[3].Value = skolkouzhemozhno[1];
            }
            sr.Close();
        } /*будет искать только в рамках [System Access] и [Event Audit]*/

        public void PoiskPoCmdAuditPol(string parametr, int Stroka, DataGridView dataGridView1)
        {
            StreamReader sr = new StreamReader(currentdirectory + "\\auditpolicy.csv", Encoding.Default);
            string dataFromFile = File.ReadAllText(currentdirectory + "\\auditpolicy.csv", Encoding.Default).Replace("\r", "");
            string pobeda = dataFromFile.Replace(" ", "");
            string[] words = pobeda.Split('\n');

            string lol = parametr.Replace(" ", "");
            string[] kek = lol.Split('|');

            foreach (string s in words)
            {
                if (Regex.IsMatch(s, kek[1]))
                    dataGridView1.Rows[Stroka].Cells[3].Value = s[s.Length - 1].ToString();
            }
            sr.Close();
        } /*будет искать только в рамках auditpolicy.csv*/

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            cvet(dataGridView1);

        } /*если мы кликнем по сортировке, то определение соответствия по цвету не пропадет*/

        private void dataGridView2_ColumnHeaderMouseClick_1(object sender, DataGridViewCellMouseEventArgs e)
        {
            cvet(dataGridView2);
        }

        private void dataGridView3_ColumnHeaderMouseClick_1(object sender, DataGridViewCellMouseEventArgs e)
        {
            cvet(dataGridView3);
        }
        
        private void dataGridView5_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            cvet(dataGridView5);
        }

        private void dataGridView4_ColumnHeaderMouseClick_1(object sender, DataGridViewCellMouseEventArgs e)
        {
            cvet(dataGridView4);
        }

        private void создатьКраткийОтчетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StreamWriter sw = new StreamWriter(currentdirectory + "\\Краткий отчет.csv", false, Encoding.Default);
            sw.Write("Имя пользователя: " + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "\t\nДата проведения проверки: " + DateTime.Now + "\t\n\t\nНомер Политики;Путь;Эталонное Значение;Текущее Значение;Статус;\t\n");
            string stroka = "";
            for (int j = 0; j < dataGridView1.RowCount; j++)
            {
                if (dataGridView1.Rows[j].Cells[4].Value.ToString() == "Не успех")
                {
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        stroka += (dataGridView1[i, j].Value.ToString() + ";");
                    }
                    stroka += "\t\n";
                }
            }
            sw.Write(stroka);
            sw.Close();

            if (File.Exists(currentdirectory + "\\Краткий отчет.csv"))
            {
                MessageBox.Show("Краткий отчет создан");
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            File.Delete(currentdirectory + "\\auditpolicy.csv");
            File.Delete(currentdirectory + "\\secpol.cfg");
        }

        private void поОператорамToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StreamWriter sw = new StreamWriter(currentdirectory + "\\Полный отчет по операторам.csv", false, Encoding.Default);
            sw.Write("Имя пользователя: " + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "\t\nДата проведения проверки: " + DateTime.Now + "\t\n\t\nНомер Политики;Путь;Эталонное Значение;Текущее Значение;Статус;\t\n");
            string stroka = "";
            for (int j = 0; j < dataGridView2.RowCount; j++)
            {
                for (int i = 0; i < dataGridView2.ColumnCount; i++)
                {
                    stroka += (dataGridView2[i, j].Value.ToString() + ";");
                }
                stroka += "\t\n";
            }
            sw.Write(stroka);
            sw.Write(label1.Text + "\t\n" + label2.Text + "\t\n" + label3.Text + "\t\n");
            sw.Close();

            if (File.Exists(currentdirectory + "\\Краткий отчет по операторам.csv"))
            {
                MessageBox.Show("Полный отчет по операторам создан");
            }
        }

        private void поКураторамИБToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StreamWriter sw = new StreamWriter(currentdirectory + "\\Полный отчет по кураторам ИБ.csv", false, Encoding.Default);
            sw.Write("Имя пользователя: " + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "\t\nДата проведения проверки: " + DateTime.Now + "\t\n\t\nНомер Политики;Путь;Эталонное Значение;Текущее Значение;Статус;\t\n");
            string stroka = "";
            for (int j = 0; j < dataGridView4.RowCount; j++)
            {
                for (int i = 0; i < dataGridView4.ColumnCount; i++)
                {
                    stroka += (dataGridView4[i, j].Value.ToString() + ";");
                }
                stroka += "\t\n";
            }
            sw.Write(stroka);
            sw.Write(label1.Text + "\t\n" + label2.Text + "\t\n" + label3.Text + "\t\n");
            sw.Close();

            if (File.Exists(currentdirectory + "\\Краткий отчет по кураторам ИБ.csv"))
            {
                MessageBox.Show("Полный отчет по кураторам ИБ создан");
            }
        }

        private void поАдмиистраторамToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StreamWriter sw = new StreamWriter(currentdirectory + "\\Полный отчет по администраторам.csv", false, Encoding.Default);
            sw.Write("Имя пользователя: " + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "\t\nДата проведения проверки: " + DateTime.Now + "\t\n\t\nНомер Политики;Путь;Эталонное Значение;Текущее Значение;Статус;\t\n");
            string stroka = "";
            for (int j = 0; j < dataGridView3.RowCount; j++)
            {
                for (int i = 0; i < dataGridView3.ColumnCount; i++)
                {
                    stroka += (dataGridView3[i, j].Value.ToString() + ";");
                }
                stroka += "\t\n";
            }
            sw.Write(stroka);
            sw.Write(label1.Text + "\t\n" + label2.Text + "\t\n" + label3.Text + "\t\n");
            sw.Close();

            if (File.Exists(currentdirectory + "\\Краткий отчет по администраторам.csv"))
            {
                MessageBox.Show("Полный отчет по администраторам создан");
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            cvet(dataGridView2);
            cvet(dataGridView3);
            cvet(dataGridView4);
            cvet(dataGridView5);
        }
    }
}

