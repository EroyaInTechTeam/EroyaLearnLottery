using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EroyaLearnLotteryV00
{
    public partial class LotteryPage : Form
    {
        public LotteryPage()
        {
            InitializeComponent();
            btn_Search.Enabled = false;
            txt_WinnerCounts.Text = "1";
        }

        #region ImportantVar
        string FilePath { get; set; }
        List<string> Members { get; set; }

        #endregion

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void LLbl_redirectToEroya_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // TNX StackOverFlow
            var ps = new ProcessStartInfo("https://www.eroyaintech.com/")
            {
                UseShellExecute = true,
                Verb = "open"
            };
            Process.Start(ps);

        }

        private void btn_dimodal_Click(object sender, EventArgs e)
        {
            //open file Dialogue
            using (OpenFileDialog OFD = new OpenFileDialog())
            {
                OFD.InitialDirectory = @"D:\projects";
                OFD.Filter = "Excel File|*.xls;*.xlsx;*.xlsm";
                OFD.RestoreDirectory = true;
                OFD.FilterIndex = 2;
                if (OFD.ShowDialog() == DialogResult.OK)
                {
                    FilePath = OFD.FileName;
                    btn_Search.Enabled = true;
                    ExcelEngine excelEngine = new ExcelEngine();
                    using (var stream = File.Open(FilePath, FileMode.Open, FileAccess.Read))
                    {
                        var excelfile = excelEngine.Excel.Workbooks.Open(stream);
                        var firstSheet = excelfile.Worksheets["Sheet1"];
                        lbl_mmbrsCount.Text = $"Members Count : {firstSheet.Rows.Count()}";
                        Members = new List<string>();
                        foreach (var item in firstSheet.Rows)
                        {
                            Members.Add(item.Columns[0].Value);
                            lb_Mmbrs.Items.Add(item.Columns[0].Value);
                        }
                    }



                }
            }
        }

        private void btn_Search_Click(object sender, EventArgs e)
        {
            int membercount = Members.Count;
            int winnerCount = 0;
            List<int> winnerIndex = new List<int>();
            if (!string.IsNullOrEmpty(txt_WinnerCounts.Text))
            {
                winnerCount = Convert.ToInt32(txt_WinnerCounts.Text);
                for (int i = 0; i < winnerCount; i++)
                {
                    Random Rand = new Random();
                    winnerIndex.Add(Rand.Next(0, membercount - 1));
                }
                string result = "";
                for (int i = 0; i < winnerIndex.Count; i++)
                {
                    result += $"Winner no.{i + 1} : InstagramID - {Members[winnerIndex[i]]}\n";
                }
                MessageBox.Show(result,"We Have Winners !",  MessageBoxButtons.OK, MessageBoxIcon.Information);




            }
            else
            {
                MessageBox.Show("Enter Winners Count","Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txt_WinnerCounts_TextChanged(object sender, EventArgs e)
        {
            int Flag = 0;
            try
            {
                if (!Int32.TryParse(txt_WinnerCounts.Text, out Flag))
                {
                    
                    txt_WinnerCounts.Text = "1";
                }
                else
                {
                    if (Flag <= 0) 
                    {
                        txt_WinnerCounts.Text = "1";
                    }
                }
            }
            catch (Exception ex)
            {
                txt_WinnerCounts.Text = "1";
            }
        }
    }
}
