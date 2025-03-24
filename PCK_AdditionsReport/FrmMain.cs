using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using Peleg.PrintJobs;

namespace Peleg.PCK_AdditionsReport
{
    public partial class FrmMain : Form
    {
        private string _srvList;
        private readonly SqlConnection _cn;
        private SqlDataAdapter _daDestinations;
        private DataTable _dtDestinations;

        private SqlDataAdapter _daSuppliers;
        private DataTable _dtSuppliers;

        private SqlDataAdapter _daServices;
        private DataTable _dtServices;



        // --------------------------------------------------------- //

        public FrmMain()
        {
            InitializeComponent();
        }

        public FrmMain(SqlConnection sqlConnection)
        {
            InitializeComponent();
            _cn = sqlConnection;
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {

            ugMain.DataSource = Tools.GetTable(PLG_GetGeneralProductsList());

            // Fill combo supplier
            _dtSuppliers = new DataTable();
            _daSuppliers = new SqlDataAdapter();

            SqlCommand cmdSqlSuppliers = new SqlCommand
            {
                Connection = _cn,
                CommandText = "PLG_GetListOfSuppliersInPackages",
                CommandType = CommandType.StoredProcedure
            };

            _daSuppliers.SelectCommand = cmdSqlSuppliers;
            _daSuppliers.Fill(_dtSuppliers);
            ucSupplier.DisplayMember = "Name";
            ucSupplier.ValueMember = "Code";
            ucSupplier.DataSource = _dtSuppliers;

            // fill combo Destinations
            _dtDestinations = new DataTable();
            _daDestinations = new SqlDataAdapter();

            SqlCommand cmdSqlDestinations = new SqlCommand
            {
                Connection = _cn,
                CommandText = "PLG_GetListOfDestinationsInPackages",
                CommandType = CommandType.StoredProcedure
            };

            _daDestinations.SelectCommand = cmdSqlDestinations;
            _daDestinations.Fill(_dtDestinations);
            ucDestination.DisplayMember = "Name";
            ucDestination.ValueMember = "Code";
            ucDestination.DataSource = _dtDestinations;

            //// fill combo services

            _dtServices = new DataTable();
            _daServices = new SqlDataAdapter();

            SqlCommand cmdSqlServices = new SqlCommand
            {
                Connection = _cn,
                CommandText = "PLG_GetListOfServiceInPackages",
                CommandType = CommandType.StoredProcedure
            };

            _daServices.SelectCommand = cmdSqlServices;
            _daServices.Fill(_dtServices);
            ucService.DisplayMember = "Name";
            ucService.ValueMember = "Code";
            ucService.DataSource = _dtServices;

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            string connstr = UP.Connection.SqlConnectionString;
            string supplier = "";
            string destination = "";
            int service;
            string company = UP.Company.Code;
            string reportPath = Application.StartupPath + @"\PLG_Report\PackageAdditions.rpt";
            int withMoney = 0;
            int withHotelName = 0;
            int option;
            DateTime fDate = dtFrom.Value;
            DateTime tDate = dtTo.Value;
            string services;

            //=================================

            if (ucSupplier.Text != "")
                supplier = ucSupplier.Value.ToString();
            if (ucDestination.Text != "")
                destination = ucDestination.Value.ToString();

            // todo add service

            if (rbPck.Checked)
                option = 0;
            else
                if (rbBoth.Checked)
                option = 1;
            else
                option = 2;

            if (chWithPrice.Checked)
                withMoney = 1;

            if (chHotelName.Checked)
                withHotelName = 1;

            if (option == 0)
                services = "";
            else
            {
                services = ugService.Rows.Where(x => (bool)x.Cells["Choose"].Value)
                    .Select(x => (string)x.Cells["Code"].Value)
                    .DefaultIfEmpty()
                    .Aggregate((current, next) => current + "," + next) ?? "";
            }

            if (option == 2)
                service = 0;
            else
            {
                if (ucService.Text == "")
                    service = 0;
                else
                    service = (int)ucService.Value;
            }

            CRManager rep = new CRManager();
            rep.Parameters.Add("FDate", fDate);
            rep.Parameters.Add("TDate", tDate);
            rep.Parameters.Add("Supplier", supplier);
            rep.Parameters.Add("Destination", destination);
            rep.Parameters.Add("WithMoney", withMoney);
            rep.Parameters.Add("Company", company);
            rep.Parameters.Add("Service", service);
            rep.Parameters.Add("Services", services);
            rep.Parameters.Add("Option", option);
            rep.Parameters.Add("WithHotelName", withHotelName);
            rep.ShowReportByName(connstr, reportPath, CRManager.ViewStyleEnum.Full);

        }

        private void ucSupplier_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.Override.HeaderAppearance.TextHAlign = HAlign.Center;
            ColumnsCollection cols = e.Layout.Bands[0].Columns;

            cols["Code"].Width = 55;
            cols["Name"].Width = 220;
        }

        private void ucDestination_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.Override.HeaderAppearance.TextHAlign = HAlign.Center;
            ColumnsCollection cols = e.Layout.Bands[0].Columns;

            cols["Code"].Width = 55;
            cols["Name"].Width = 220;

        }

        //################# SQL PART ###############
        SqlCommand PLG_GetGeneralProductsList()
        {
            SqlCommand cmd = _cn.CreateCommand();
            cmd.CommandText = "PLG_GetGeneralProductsList";
            cmd.CommandType = CommandType.StoredProcedure;
            return cmd;
        }

        SqlCommand PLG_GetProductsList_ByGeneral(string ms)
        {
            SqlCommand cmd = _cn.CreateCommand();
            cmd.CommandText = "PLG_GetProductsList_ByGeneral";
            cmd.CommandType = CommandType.StoredProcedure;

            cmd.Parameters.Add("@MS", SqlDbType.VarChar, 300);
            cmd.Parameters["@MS"].Value = ms;
            return cmd;
        }

        //#######################################################
        private void ugMain_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.Override.HeaderAppearance.TextHAlign = HAlign.Center;
            ColumnsCollection cols = e.Layout.Bands[0].Columns;
            cols["Choose"].Width = 55;
            cols["Code"].Width = 45;
            cols["Name"].Width = 150;
            cols["Code"].CellClickAction = CellClickAction.RowSelect;
            cols["Name"].CellClickAction = CellClickAction.RowSelect;
            cols["Choose"].CellClickAction = CellClickAction.Edit;

        }

        private void ugService_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.Override.HeaderAppearance.TextHAlign = HAlign.Center;
            ColumnsCollection cols = e.Layout.Bands[0].Columns;
            cols["Choose"].Width = 55;
            cols["Code"].Width = 45;
            cols["Name"].Width = 150;
            cols["Code"].CellClickAction = CellClickAction.RowSelect;
            cols["Name"].CellClickAction = CellClickAction.RowSelect;
            cols["Choose"].CellClickAction = CellClickAction.Edit;

        }

        private void ugService_CellChange(object sender, CellEventArgs e)
        {

        }

        private void ugMain_CellChange(object sender, CellEventArgs e)
        {
            e.Cell.Row.Update();
            RefreshServiceList();

        }

        private void RefreshServiceList()
        {
            _srvList = ugMain.Rows.Where(x => (bool)x.Cells["Choose"].Value)
                      .Select(x => (string)x.Cells["Code"].Value)
                      .DefaultIfEmpty()
                      .Aggregate((current, next) => current + "," + next);

            ugService.DataSource = _srvList == null ? null : Tools.GetTable(PLG_GetProductsList_ByGeneral(_srvList));
        }

        private void chAllMain_CheckedChanged(object sender, EventArgs e)
        {
            foreach (UltraGridRow row in ugMain.Rows)
                row.Cells["Choose"].Value = chAllMain.Checked;

            RefreshServiceList();
        }

        private void chAllServices_CheckedChanged(object sender, EventArgs e)
        {
            foreach (UltraGridRow row in ugService.Rows)
                row.Cells["Choose"].Value = chAllServices.Checked;

        }

        private void chSelectedMain_CheckedChanged(object sender, EventArgs e)
        {
            if (chSelectedMain.Checked)
                foreach (UltraGridRow row in ugMain.Rows.Where(x => !(bool)x.Cells["Choose"].Value))
                    row.Hidden = true;
            else
                foreach (UltraGridRow row in ugMain.Rows)
                    row.Hidden = false;
        }

        private void chSelectedServices_CheckedChanged(object sender, EventArgs e)
        {
            if (chSelectedServices.Checked)
                foreach (UltraGridRow row in ugService.Rows.Where(x => !(bool)x.Cells["Choose"].Value))
                    row.Hidden = true;
            else
                foreach (UltraGridRow row in ugService.Rows)
                    row.Hidden = false;
        }

        private void ucService_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.Override.HeaderAppearance.TextHAlign = HAlign.Center;
            ColumnsCollection cols = e.Layout.Bands[0].Columns;

            cols["Code"].Width = 55;
            cols["Name"].Width = 220;

        }

        private void btnExcel_Click(object sender, EventArgs e)
        {

            int withPrice = 0;
            int withHotel = 0;

            string clerk = UP.User.Name;
            var pathDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string fileName = "PackageAdd_" + clerk + "_" + DateTime.Now.ToString("ddMMyy") + "_" + DateTime.Now.ToString("HHmm") + ".xlsx";

            string supplier = "";
            string destination = "";
            int service;
            int option;
            DateTime fDate = dtFrom.Value;
            DateTime tDate = dtTo.Value;
            string services;

            //=================================
            if (chWithPrice.Checked)
                withPrice = 1;

            if (chHotelName.Checked)
                withHotel = 1;

            if (ucSupplier.Text != "")
                supplier = ucSupplier.Value.ToString();
            if (ucDestination.Text != "")
                destination = ucDestination.Value.ToString();

            // todo add service

            if (rbPck.Checked)
                option = 0;
            else
                if (rbBoth.Checked)
                option = 1;
            else
                option = 2;

            if (option == 0)
                services = "";
            else
            {
                services = ugService.Rows.Where(x => (bool)x.Cells["Choose"].Value)
                    .Select(x => (string)x.Cells["Code"].Value)
                    .DefaultIfEmpty()
                    .Aggregate((current, next) => current + "," + next) ?? "";
            }

            if (option == 2)
                service = 0;
            else
            {
                if (ucService.Text == "")
                    service = 0;
                else
                    service = (int)ucService.Value;
            }

            var ed = new ExportData2Excel.Work();
            ed.ExportSp(_cn.ConnectionString, fileName, "PLG_PCK_ReportAdditions4Excel", new object[] { fDate, tDate, supplier, destination, service, services, option, withPrice, withHotel });

        }
    }
}
