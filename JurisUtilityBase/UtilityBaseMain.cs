using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public int FldClient { get; set; }

        public int FldMatter { get; set; }

        private string BnkCode = "";

        private int totalRecRows = 0;

        private bool validRec = false;

        private string lastReconBal = "";

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }


            string bankAcct;
            comboBox1.ClearItems();
            string SQLTkpr = "SELECT BnkCode + '  ' + BnkDesc as bankaccount FROM BankAccount";
            DataSet myRSTkpr = _jurisUtility.RecordsetFromSQL(SQLTkpr);

            if (myRSTkpr.Tables[0].Rows.Count == 0)
                comboBox1.SelectedIndex = -1;
            else
            {
                foreach (DataTable table in myRSTkpr.Tables)
                {

                    foreach (DataRow dr in table.Rows)
                    {
                        bankAcct = dr["bankaccount"].ToString();
                        comboBox1.Items.Add(bankAcct);
                    }
                }

            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            // Enter your SQL code here
            // To run a T-SQL statement with no results, int RecordsAffected = _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            // To get an ADODB.Recordset, ADODB.Recordset myRS = _jurisUtility.RecordsetFromSQL(SQL);

            if (validRec) //do nothing if a valid rec wasnt found on the bank account
            {
                string SQL = "SELECT BRHStmtDate, BRHBookLastStmtDate FROM BankReconHistory WHERE (BRHBank = '" + BnkCode + "') AND (BRHRecorded = 'N')";
                DataSet rectest1 = _jurisUtility.RecordsetFromSQL(SQL);
                if (rectest1.Tables[0].Rows.Count > 0)  //there ARE unrecorded recs
                {
                    DialogResult inProgress = MessageBox.Show("There is a bank reconciliation in progress for Bank " + BnkCode + ". That reconciliation  will be deleted." + "\r\n" + "The last recorded reconciliation (statement date: " + textBoxStateDate + ") will be set to 'unrecorded'.  Do you wish to continue?", "Reconciliation in progress", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (inProgress == System.Windows.Forms.DialogResult.Yes)
                    {

                        /////////delete unrecorded items
                        SQL = "DELETE FROM BankReconHistory WHERE (BRHBank = '" + BnkCode + "') AND (BRHRecorded = 'N')";
                        _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                        SQL = "UPDATE CheckRegister SET CkRegCleared = 'N' WHERE (CkRegBank = '" + BnkCode + "') AND (CkRegCleared = 'P')";
                        _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                        /////////set last rec to unrecorded after unrecorded rec was removed
                        setBankRecToUnrecorded();
                        updateBankAccount();
                    }
                    else
                        MessageBox.Show("No changes were made");
                }
                else  //there are NO unrecorded recs
                {
                    setBankRecToUnrecorded();
                    updateBankAccount();


                }

                UpdateStatus("Bank Rec undone.", 1, 1);

                MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);





            }
            validRec = false;
        }

        private void setBankRecToUnrecorded()
        {
            string SQL = "UPDATE BankReconHistory SET BRHRecorded = 'N', BRHRecordedDate = CONVERT(DATETIME, '12/31/1899 00:00:00', 101) " +
                "WHERE (BRHRecorded = 'Y') AND (BRHBank = '" + BnkCode + "') AND (BRHRecordedDate = '" + textBoxRecDate.Text + "') AND " +
                "(BRHStmtDate = '" + textBoxStateDate.Text + "')";
            _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            SQL = "UPDATE CheckRegister SET CkRegCleared = 'P' WHERE (CkRegBank = '" + BnkCode + "') AND (CkRegCleared = 'Y') AND (CkRegReconDate = '" 
                + textBoxStateDate.Text + "')";
            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

        }


        private void updateBankAccount()
        {
            if (totalRecRows == 1)
            {
                string SQL = "UPDATE BankAccount SET BnkLastReconDate = CONVERT(DATETIME, '12/31/1899 00:00:00', 101), BnkLastReconBal = 0 WHERE BnkCode = '" + BnkCode + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            }
            else
            {
                string SQL = "UPDATE BankAccount SET BnkLastReconDate = '" + textBoxRecDate.Text + "', BnkLastReconBal = " + lastReconBal + " WHERE BnkCode = '" + BnkCode + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            }


        }


        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {

            System.Environment.Exit(0);
          
        }


        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex > -1)
            {
                BnkCode = comboBox1.Text.Split(' ')[0];
                string SQL = "select BRHRecordedDate, BRHStmtDate, BRHStmtOpenBal, BRHStmtEndBal, BRHBookClearedBal " +
                    "FROM         BankReconHistory " +
                    "WHERE     (BRHRecorded = 'Y') AND (BRHBank = '" + BnkCode + "') " +
                    "ORDER BY BRHStmtDate DESC, BRHRecordedDate DESC";

                DataSet rec1 = _jurisUtility.RecordsetFromSQL(SQL);
                if (rec1.Tables[0].Rows.Count > 0)
                {
                    textBoxBegBal.Text = double.Parse(rec1.Tables[0].Rows[0][2].ToString()).ToString("#.##");
                    textBoxEndBal.Text = double.Parse(rec1.Tables[0].Rows[0][3].ToString()).ToString("#.##");
                    textBoxRecDate.Text = Convert.ToDateTime(rec1.Tables[0].Rows[0][0].ToString()).ToString("MM/dd/yyyy");
                    textBoxStateDate.Text = Convert.ToDateTime(rec1.Tables[0].Rows[0][1].ToString()).ToString("MM/dd/yyyy");
                    lastReconBal = double.Parse(rec1.Tables[0].Rows[0][3].ToString()).ToString("#.##");
                    totalRecRows = rec1.Tables[0].Rows.Count;
                    validRec = true;
                }
                else
                {
                    MessageBox.Show("There are no reconciliations for that bank account", "Input error");
                    validRec = false;
                }
                rec1.Clear();
            }
        }


    }
}
