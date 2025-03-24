using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Peleg.PCK_AdditionsReport
{
    public class Manager : StartMenuInterface.IStartMenu
    {
        public void Start(string connStr)
        {
            SqlConnection cn = OpenConnection(connStr);
            if (cn != null)
                new FrmMain(cn).ShowDialog();
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

        #region _IStartMenu Members

        public void Run(ref string ConnectionString, ref string TaskName)
        {
            Start(NaumTools.Utils.Ado2Sql(ConnectionString));

        }

        #endregion

    }
}
