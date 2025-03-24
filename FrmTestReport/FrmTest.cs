using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FrmTestReport
{
    public partial class FrmTest : Form
    {
        private string connString;

        public FrmTest()
        {
            InitializeComponent();
        }

        private static SqlConnection OpenConnection(string connStr)
        {
            try
            {
                SqlConnection cn = new SqlConnection(connStr);
                cn.Open();
                return cn;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();

        }

        private void FrmTest_Load(object sender, EventArgs e)
        {
            connString = UP.Connection.SqlConnectionString;
            SqlConnection _cn = OpenConnection(connString);

        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            Peleg.PCK_AdditionsReport.Manager EAF = new Peleg.PCK_AdditionsReport.Manager();
            EAF.Start(connString);

        }
    }
}
