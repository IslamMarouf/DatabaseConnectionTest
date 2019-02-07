using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using DatabaseConnectionTest.Properties;

namespace DatabaseConnectionTest
{
    public partial class Form1 : Form {
        public string subjectsDbPath;
        public readonly OleDbConnection SubjectsConnection = new OleDbConnection();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e) {
            subjectsDbPath = Settings.Default.subjectsDbPath;
            subjectsDbPath = "";
            

            

            if (subjectsDbPath.Equals("")) {
                subjectsDbPath = GetDatabasePath();
            }
            else {
                /*
              This is an encrypted access 2016 database. 
              To encrypt data in 2016 do the follow.
              Under File - Options - Client Settings (scroll to the bottom)...
              Default Open Mode = Shared
              Default record locking = No Locks
              Encryption Method = Use legacy
              Unencrypt and re-encrypt the database.
            */
                string strSubjectsCon = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                                        @"Data Source=" + subjectsDbPath + "Subjects.accdb;" +
                                        @"Jet OLEDB:Database Password=0163575879";

                SubjectsConnection.ConnectionString = strSubjectsCon;
            }

            try
            {
                SubjectsConnection.Open();
                MessageBox.Show("Subjects Database connection is successful.");
                MessageBox.Show(SubjectsConnection.ConnectionString);
            }
            catch (OleDbException oleDbEx)
            {
                MessageBox.Show("Access Error:" +
                                "\n\nError Code = " + oleDbEx.ErrorCode +
                                "\n\nError Message = " + oleDbEx.Message);
            }
            catch (InvalidOperationException invOpEx)
            {
                DialogResult =
                MessageBox.Show("", "تعذر الوصول الى قاعدة البيانات", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1);

                switch (DialogResult) {
                    case DialogResult.Retry:
                        subjectsDbPath = GetDatabasePath();
                        break;
                    case DialogResult.Cancel:
                        MessageBox.Show("Invalid Message = " +
                                        "\n\nError Message = " + invOpEx.Message);
                        break;
                }       
            }

            if (SubjectsConnection.State != ConnectionState.Open) {
                simpleButton1.Enabled = false;
                simpleButton2.Enabled = false;
                textEdit1.Enabled = false;}
        }

        private string GetDatabasePath() {
            string subjectsDbPath = string.Empty;DialogResult dlgresult;
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            // Getting installation folder from user
            folderDlg.Description = "برجاء اختيار مجلد لحفظ قاعدة البيانات";
            folderDlg.RootFolder = Environment.SpecialFolder.MyComputer;

            dlgresult = folderDlg.ShowDialog();

            if (dlgresult == DialogResult.OK) {
                // Assigning operators
                subjectsDbPath = folderDlg.SelectedPath;
                File.WriteAllBytes(subjectsDbPath + @"\Subjects.accdb", Properties.Resources.Subjects);
                Settings.Default.subjectsDbPath = subjectsDbPath + @"\";
                Settings.Default.Save();
                MessageBox.Show(subjectsDbPath);
            }

            return subjectsDbPath;
        }
    }
}
