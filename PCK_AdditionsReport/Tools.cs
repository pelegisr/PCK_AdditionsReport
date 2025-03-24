using System.Data;
using System.Data.SqlClient;

namespace Peleg.PCK_AdditionsReport
{
    internal class Tools
    {

        public static DataTable GetTable(SqlCommand command)
        {
            DataTable dt = new DataTable();
            new SqlDataAdapter(command).Fill(dt);
            return dt;
        }

    }
}
