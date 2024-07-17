using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Web;
using System.Web.Script.Serialization;
using System.Web.Services;
using System.Web.UI;
using System.Web.UI.WebControls;
using Newtonsoft;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Drawing;
namespace FinAnalyz_NetNew
{
    public partial class BSA2 : System.Web.UI.Page
    {
        private string connectionString = ConfigurationManager.ConnectionStrings["dbcs"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {

            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            if (!IsPostBack)
            {
                //BindExceptionalTransactions();
                DateTime startDate = DateTime.Now.AddMonths(-6); // Default last 6 months
                DateTime endDate = DateTime.Now;
                Bind_ddlAccountNumber();
                if (ddlAccountNumber.Items.Count > 0)
                {

                    ddlAccountNumber.SelectedIndex = 0;
                    LoadChartData(startDate, endDate);
                    BindGridView(startDate, endDate);
                    BindSummaryGridView(startDate, endDate);
                    BindMonthlySummaryGridView(startDate, endDate);
                    BindBankDetails(startDate, endDate);
                    BindMonthlyBalance(startDate, endDate);
                    //GetChartData(startDate, endDate);
                    LoadCandlestickChart(startDate, endDate);

                }
            }
        }



        private void BindBankDetails(DateTime startDate, DateTime endDate)
        {
            string accountId = ddlAccountNumber.SelectedValue;

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                // Fetch account details
                string accountQuery = @"spBSA_Savings";

                using (SqlCommand cmd = new SqlCommand(accountQuery, con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@QueryType", "AccountDetails");
                    cmd.Parameters.AddWithValue("@AccountId", accountId);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            lblAccountName.Text = reader["account_name"].ToString();
                            lblBankName.Text = reader["bank_name"].ToString();
                            lblAccountNumber.Text = reader["account_number"].ToString();
                            lblAccountType.Text = reader["account_type"].ToString();
                        }
                    }
                }

                // Fetch transaction details including balances and period
                string transactionQuery = @"spBSA_Savings";

                using (SqlCommand cmd = new SqlCommand(transactionQuery, con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@QueryType", "TransactionData");
                    cmd.Parameters.AddWithValue("@AccountId", accountId);
                    cmd.Parameters.AddWithValue("@StartDate", startDate);
                    cmd.Parameters.AddWithValue("@EndDate", endDate);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            DateTime start = reader["StartDate"] != DBNull.Value ? Convert.ToDateTime(reader["StartDate"]) : DateTime.MinValue;
                            DateTime end = reader["EndDate"] != DBNull.Value ? Convert.ToDateTime(reader["EndDate"]) : DateTime.MinValue;

                            lblPeriod.Text = start != DateTime.MinValue && end != DateTime.MinValue ? $"{start:dd-MM-yyyy} to {end:dd-MM-yyyy}" : "Not available";

                            int numberOfMonths = reader["NumberOfMonths"] != DBNull.Value ? Convert.ToInt32(reader["NumberOfMonths"]) : 0;
                            lblNumberOfMonths.Text = numberOfMonths.ToString();

                            lblOpeningBalance.Text = reader["TotalCredits"] != DBNull.Value ? reader["TotalCredits"].ToString() : "0";
                            lblClosingBalance.Text = reader["ClosingBalance"] != DBNull.Value ? reader["ClosingBalance"].ToString() : "0";

                            lblCurrency.Text = "INR"; // Assuming currency is INR, change as needed
                        }
                        else
                        {
                            lblPeriod.Text = "Not available";
                            lblNumberOfMonths.Text = "0";
                            lblOpeningBalance.Text = "0";
                            lblClosingBalance.Text = "0";
                            lblCurrency.Text = "INR"; // Assuming currency is INR, change as needed
                        }
                    }
                }
            }
        }

        private void BindMonthlySummaryGridView(DateTime startDate, DateTime endDate)
        {
            if (ddlAccountNumber.SelectedIndex == 0)
            {
                ErrorMessageLabel.Text = "Please select a valid account number.";
                return;
            }

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = @"spBSA_Savings";

                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@QueryType", "MonthlyIncomeExpense");
                    cmd.Parameters.AddWithValue("@AccountId", ddlAccountNumber.SelectedValue);
                    cmd.Parameters.AddWithValue("@StartDate", startDate);
                    cmd.Parameters.AddWithValue("@EndDate", endDate);
                    cmd.Connection = con;
                    con.Open();
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    con.Close();
                    if (dt.Rows.Count > 0)
                    {
                        MonthlySummaryGridView.DataSource = dt;
                        MonthlySummaryGridView.DataBind();
                        EODPanel.Visible = true;
                        lblEOD.Text = "";
                    }
                    else
                    {
                        EODPanel.Visible = false;
                        lblEOD.Text = "No Data Available...";
                    }
                }
            }
        }
        private void BindSummaryGridView(DateTime startDate, DateTime endDate)
        {
            if (ddlAccountNumber.SelectedIndex == 0)
            {
                ErrorMessageLabel.Text = "Please select a valid account number.";
                return;
            }

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = @"spBSA_Savings";

                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@QueryType", "TotalIncomeExpense");
                    cmd.Parameters.AddWithValue("@AccountId", ddlAccountNumber.SelectedValue);
                    cmd.Parameters.AddWithValue("@StartDate", startDate);
                    cmd.Parameters.AddWithValue("@EndDate", endDate);
                    cmd.Connection = con;
                    con.Open();
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    con.Close();
                    if (dt.Rows.Count > 0)
                    {
                        SummaryGridView.DataSource = dt;
                        SummaryGridView.DataBind();

                    }

                }
            }
        }
        void Bind_ddlAccountNumber()
        {

            SqlConnection con = new SqlConnection(connectionString);
            string query = "select bad_id, account_number from bank_account_details";
            SqlDataAdapter sda = new SqlDataAdapter(query, con);
            DataTable data = new DataTable();
            sda.Fill(data);
            ddlAccountNumber.DataSource = data;
            ddlAccountNumber.DataTextField = "account_number";
            ddlAccountNumber.DataValueField = "bad_id";
            ddlAccountNumber.DataBind();

            ListItem Select_Item = new ListItem("-- Select Account Number --", "0");

            Select_Item.Selected = true;
            ddlAccountNumber.Items.Insert(0, Select_Item);
        }

        protected void ddlTimeFrame_SelectedIndexChanged(object sender, EventArgs e)
        {

            txtStartDate.Text = string.Empty;
            txtEndDate.Text = string.Empty;
            DateTime startDate = DateTime.MinValue;
            DateTime endDate = DateTime.MaxValue;

            string selectedTimeFrame = ddlTimeFrame.SelectedValue;

            if (selectedTimeFrame == "0")
            {
                // Show message or popup indicating that data is not available
                ClientScript.RegisterStartupScript(this.GetType(), "alert", "alert('Data not available. Please select a valid date range.');", true);
                return;
            }
            switch (selectedTimeFrame)
            {
                case "today":
                    startDate = DateTime.Today;
                    endDate = DateTime.Today;
                    break;
                case "yesterday":
                    startDate = DateTime.Today.AddDays(-1);
                    endDate = DateTime.Today.AddDays(-1);
                    break;
                case "thisWeek":
                    startDate = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);
                    endDate = DateTime.Today;
                    break;
                case "lastWeek":
                    startDate = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek - 7);
                    endDate = startDate.AddDays(6);
                    break;
                case "thisMonth":
                    startDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                    endDate = startDate.AddMonths(1).AddDays(-1);
                    break;
                case "lastMonth":
                    startDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(-1);
                    endDate = startDate.AddMonths(1).AddDays(-1);
                    break;
                case "last3Months":
                    startDate = DateTime.Today.AddMonths(-3);
                    endDate = DateTime.Today;
                    break;
                case "last6Months":
                    startDate = DateTime.Today.AddMonths(-6);
                    endDate = DateTime.Today;
                    break;
                case "last12Months":
                    startDate = DateTime.Today.AddMonths(-12);
                    endDate = DateTime.Today;
                    break;
                case "lastFinancialYear":
                    if (DateTime.Now.Month >= 4) // Current month is April or later
                    {
                        startDate = new DateTime(DateTime.Now.Year - 1, 4, 1);
                        endDate = new DateTime(DateTime.Now.Year, 3, 31);
                    }
                    else // Current month is before April
                    {
                        startDate = new DateTime(DateTime.Now.Year - 2, 4, 1);
                        endDate = new DateTime(DateTime.Now.Year - 1, 3, 31);
                    }
                    break;
            }

            LoadChartData(startDate, endDate);
            BindSummaryGridView(startDate, endDate);
            BindMonthlySummaryGridView(startDate, endDate);
            BindBankDetails(startDate, endDate);
            BindGridView(startDate, endDate);
            BindMonthlyBalance(startDate, endDate);
            LoadCandlestickChart(startDate, endDate);


        }


        protected void SearchBtn_Click(object sender, EventArgs e)
        {
            string startDateString = txtStartDate.Text.Trim();
            string endDateString = txtEndDate.Text.Trim();
            DateTime startDate;
            DateTime endDate;

            if (DateTime.TryParseExact(startDateString, "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out startDate) &&
                DateTime.TryParseExact(endDateString, "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out endDate))
            {
                LoadChartData(startDate, endDate);
                BindGridView(startDate, endDate);
                BindSummaryGridView(startDate, endDate); // Bind summary grid view
                BindMonthlySummaryGridView(startDate, endDate);
                BindBankDetails(startDate, endDate);
                BindMonthlyBalance(startDate, endDate);
                LoadCandlestickChart(startDate, endDate);
                ErrorMessageLabel.Text = string.Empty;
            }
            else
            {
                startDate = DateTime.Now.AddMonths(-6);
                endDate = DateTime.Now;
                LoadChartData(startDate, endDate);
                BindGridView(startDate, endDate);
                BindSummaryGridView(startDate, endDate);
                BindMonthlySummaryGridView(startDate, endDate);
                BindMonthlyBalance(startDate, endDate);
                BindBankDetails(startDate, endDate);
                LoadCandlestickChart(startDate, endDate);
                ErrorMessageLabel.Text = "Invalid date format. Please enter dates in DD-MM-YYYY format.";
            }
        }
        private void LoadChartData(DateTime startDate, DateTime endDate)
        {
            if (ddlAccountNumber.SelectedIndex == 0)
            {
                ErrorMessageLabel.Text = "Please select a valid account number.";
                return;
            }

            DataTable dt = GetIncomeExpenseData(startDate, endDate);

            var groupedData = dt.AsEnumerable()
                                .GroupBy(r => new { Month = r.Field<DateTime>("value_date").ToString("yyyy-MM") })
                                .Select(g => new
                                {
                                    Month = g.Key.Month,
                                    Credits = g.Sum(r => r.Field<decimal>("credits")),
                                    Debits = g.Sum(r => r.Field<decimal>("debits"))
                                })
                                .OrderBy(g => g.Month)
                                .ToList();

            string dates = "[";
            string credits = "[";
            string debits = "[";

            foreach (var data in groupedData)
            {
                dates += "\"" + data.Month + "\",";
                credits += data.Credits + ",";
                debits += data.Debits + ",";
            }

            dates = dates.TrimEnd(',') + "]";
            credits = credits.TrimEnd(',') + "]";
            debits = debits.TrimEnd(',') + "]";
            ClientScript.RegisterStartupScript(this.GetType(), "chartScript", $"drawChart({dates}, {credits}, {debits});", true);
        }

        private DataTable GetIncomeExpenseData(DateTime startDate, DateTime endDate)
        {
            DataTable dt = new DataTable();

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(@"spBSA_Savings", con))
                {

                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@QueryType", "TransactionDetails");
                    cmd.Parameters.AddWithValue("@StartDate", startDate);
                    cmd.Parameters.AddWithValue("@EndDate", endDate);
                    cmd.Parameters.AddWithValue("@AccountId", ddlAccountNumber.SelectedValue);

                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(dt);
                    }
                }
            }

            return dt;
        }

        private void BindGridView(DateTime startDate, DateTime endDate)
        {
            if (ddlAccountNumber.SelectedIndex == 0)
            {
                ErrorMessageLabel.Text = "Please select a valid account number.";
                return;
            }

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = @"spBSA_Savings";

                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@QueryType", "AllTransactions");
                    cmd.Parameters.AddWithValue("@AccountId", ddlAccountNumber.SelectedItem.Value);
                    cmd.Parameters.AddWithValue("@StartDate", startDate);
                    cmd.Parameters.AddWithValue("@EndDate", endDate);
                    cmd.Connection = con;
                    con.Open();
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    if (dt.Rows.Count > 0)
                    {
                        GridView1.DataSource = dt;
                        GridView1.DataBind();
                        lblTransactionDetails.Text = "";

                    }
                    else
                    {
                        GridView1.DataSource = "";
                        GridView1.DataBind();
                        lblTransactionDetails.Text = "No Data Available...";

                    }
                    con.Close();
                }
            }
        }

        protected void MonthlySummaryGridView_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                // Assuming "Month" is the first column (index 0) in your GridView
                string originalDate = e.Row.Cells[0].Text;
                if (!string.IsNullOrEmpty(originalDate))
                {
                    DateTime parsedDate;
                    if (DateTime.TryParseExact(originalDate, "yyyy-MM", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                    {
                        e.Row.Cells[0].Text = parsedDate.ToString("MMMM yyyy");
                    }
                }
            }
        }

        protected void ddlAccountNumber_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddlAccountNumber.SelectedValue == "0")
            {
                ddlTimeFrame.ClearSelection();
                txtStartDate.Text = string.Empty;
                txtEndDate.Text = string.Empty;
                ddlTimeFrame.Enabled = false;
                txtStartDate.Enabled = false;
                txtEndDate.Enabled = false;
                SearchBtn.Enabled = false;
                ErrorMessageLabel.Text = "Please select a valid account number.";
                return;
            }

            DateTime startDate = DateTime.Now.AddMonths(-6); // Default last 6 months
            DateTime endDate = DateTime.Now;
            LoadChartData(startDate, endDate);
            BindSummaryGridView(startDate, endDate);
            BindMonthlySummaryGridView(startDate, endDate);
            BindGridView(startDate, endDate);
            BindBankDetails(startDate, endDate);
            BindMonthlyBalance(startDate, endDate);
            //GetChartData(startDate, endDate);
            LoadCandlestickChart(startDate, endDate);
            LoanTransactionSummary(startDate, endDate);
            LoanTransactionDetails(startDate, endDate);
            ErrorMessageLabel.Text = "";
            BankDetailsPanel.Visible = true;
            panelHide.Visible = false;
            hidetxt.Visible = false;
            txtEOD.Visible = false;
            txtRemBal.Visible = false;
            txtHighestTransax.Visible = false;
            txtChequeRtnTransax.Visible = false;
            txtLoanSheet.Visible = false;
            ddlTimeFrame.Enabled = true;
            txtStartDate.Enabled = true;
            txtEndDate.Enabled = true;
            SearchBtn.Enabled = true;

        }

        protected void btnDownloadExcel_Click(object sender, EventArgs e)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                // Create worksheets
                var bankDetailsWorksheet = excel.Workbook.Worksheets.Add("BankDetails");
                var transactionDetailsWorksheet = excel.Workbook.Worksheets.Add("TransactionDetails");
                var eodAnalysisWorksheet = excel.Workbook.Worksheets.Add("EODAnalysis");
                var chartDataWorksheet = excel.Workbook.Worksheets.Add("ChartData");


                // Load data into worksheets
                LoadBankDetails(bankDetailsWorksheet);
                LoadTransactionDetails(transactionDetailsWorksheet);
                LoadEODAnalysis(eodAnalysisWorksheet);
                //LoadChartData(chartDataWorksheet);
                LoadGridViewDataAndChart(chartDataWorksheet);

                // Write the Excel file to response
                WriteExcelToResponse(excel, "PageDataWithChart.xlsx");
            }
        }

        private void LoadBankDetails(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A1"].Value = "Bank Name";
            worksheet.Cells["B1"].Value = lblBankName.Text;

            worksheet.Cells["A2"].Value = "Account No.";
            worksheet.Cells["B2"].Value = lblAccountNumber.Text;

            worksheet.Cells["A3"].Value = "Period";
            worksheet.Cells["B3"].Value = lblPeriod.Text;

            worksheet.Cells["A4"].Value = "No. of Months";
            worksheet.Cells["B4"].Value = lblNumberOfMonths.Text;

            worksheet.Cells["A5"].Value = "Account Type";
            worksheet.Cells["B5"].Value = lblAccountType.Text;

            worksheet.Cells["A6"].Value = "Currency";
            worksheet.Cells["B6"].Value = lblCurrency.Text;

            worksheet.Cells["A7"].Value = "Opening Balance";
            worksheet.Cells["B7"].Value = lblOpeningBalance.Text;

            worksheet.Cells["A8"].Value = "Closing Balance";
            worksheet.Cells["B8"].Value = lblClosingBalance.Text;

            // Apply styling
            for (int i = 1; i <= 8; i++)
            {
                worksheet.Cells[i, 1].Style.Font.Bold = true;
                worksheet.Cells[i, 1, i, 2].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                worksheet.Cells[i, 1, i, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[i, 1, i, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                worksheet.Cells[i, 1, i, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            }
        }

        private void LoadTransactionDetails(ExcelWorksheet worksheet)
        {
            // Headers
            worksheet.Cells["A1"].Value = "ID";
            worksheet.Cells["B1"].Value = "Transaction Type";
            worksheet.Cells["C1"].Value = "Debits";
            worksheet.Cells["D1"].Value = "Credits";
            worksheet.Cells["E1"].Value = "Current Balance";
            worksheet.Cells["F1"].Value = "Transaction Time";
            worksheet.Cells["G1"].Value = "Date";

            // Apply header styling
            using (var range = worksheet.Cells["A1:G1"])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                range.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
            }

            // Get GridView data
            if (GridView1 != null && GridView1.Rows.Count > 0)
            {
                for (int i = 0; i < GridView1.Rows.Count; i++)
                {
                    GridViewRow row = GridView1.Rows[i];
                    Label lblID = (Label)row.FindControl("lblID");
                    Label lblTransactionType = (Label)row.FindControl("lblTransactionType");
                    Label lblDebits = (Label)row.FindControl("lblDebits");
                    Label lblCredits = (Label)row.FindControl("lblCredits");
                    Label lblCurrentBalance = (Label)row.FindControl("lblCurrentBalance");
                    Label lblTransactionTime = (Label)row.FindControl("lblTransactionTime");
                    Label lblDate = (Label)row.FindControl("lblDate");

                    // Ensure controls are found before accessing their Text property
                    if (lblID != null && lblTransactionType != null && lblDebits != null &&
                        lblCredits != null && lblCurrentBalance != null && lblTransactionTime != null &&
                        lblDate != null)
                    {
                        worksheet.Cells[i + 2, 1].Value = lblID.Text;
                        worksheet.Cells[i + 2, 2].Value = lblTransactionType.Text;
                        worksheet.Cells[i + 2, 3].Value = lblDebits.Text;
                        worksheet.Cells[i + 2, 4].Value = lblCredits.Text;
                        worksheet.Cells[i + 2, 5].Value = lblCurrentBalance.Text;
                        worksheet.Cells[i + 2, 6].Value = lblTransactionTime.Text;
                        worksheet.Cells[i + 2, 7].Value = lblDate.Text;

                        // Apply cell styling
                        for (int j = 1; j <= 7; j++)
                        {
                            worksheet.Cells[i + 2, j].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 2, j].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        }
                    }
                }

               
            }
            else
            {
                throw new Exception("GridView1 is either null or has no rows.");
            }
        }

        private void LoadEODAnalysis(ExcelWorksheet worksheet)
        {
            // Headers
            worksheet.Cells["A1"].Value = "Month";
            worksheet.Cells["B1"].Value = "Monthly Income";
            worksheet.Cells["C1"].Value = "Monthly Expense";

            // Apply header styling
            using (var range = worksheet.Cells["A1:C1"])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                range.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
            }

            // Data rows
            for (int i = 0; i < MonthlySummaryGridView.Rows.Count; i++)
            {
                GridViewRow row = MonthlySummaryGridView.Rows[i];
                worksheet.Cells[i + 2, 1].Value = row.Cells[0].Text;
                worksheet.Cells[i + 2, 2].Value = row.Cells[1].Text;
                worksheet.Cells[i + 2, 3].Value = row.Cells[2].Text;

                // Apply cell styling
                for (int j = 1; j <= 3; j++)
                {
                    worksheet.Cells[i + 2, j].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells[i + 2, j].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                }
            }

           

            // Chart
            var chart = worksheet.Drawings.AddChart("IncomeExpenseChart", eChartType.ColumnClustered);
            chart.Title.Text = "Income and Expense Over Time";
            chart.SetPosition(5, 0, 4, 0);
            chart.SetSize(800, 400);

            var incomeSeries = chart.Series.Add(worksheet.Cells[2, 2, MonthlySummaryGridView.Rows.Count + 1, 2], worksheet.Cells[2, 1, MonthlySummaryGridView.Rows.Count + 1, 1]);
            incomeSeries.Header = "Monthly Income";

            var expenseSeries = chart.Series.Add(worksheet.Cells[2, 3, MonthlySummaryGridView.Rows.Count + 1, 3], worksheet.Cells[2, 1, MonthlySummaryGridView.Rows.Count + 1, 1]);
            expenseSeries.Header = "Monthly Expense";

            int summaryStartRow = MonthlySummaryGridView.Rows.Count + 3;
            worksheet.Cells[summaryStartRow, 1].Value = "Total Income";
            worksheet.Cells[summaryStartRow, 2].Value = "Total Expense";
            worksheet.Cells[summaryStartRow, 3].Value = "Total Balance";

            for (int i = 0; i < SummaryGridView.Rows.Count; i++)
            {
                GridViewRow row = SummaryGridView.Rows[i];
                worksheet.Cells[summaryStartRow + 1 + i, 1].Value = row.Cells[0].Text;
                worksheet.Cells[summaryStartRow + 1 + i, 2].Value = row.Cells[1].Text;
                worksheet.Cells[summaryStartRow + 1 + i, 3].Value = row.Cells[2].Text;
            }

            // Formatting summary rows
            using (var range = worksheet.Cells[summaryStartRow + 1, 1, summaryStartRow + SummaryGridView.Rows.Count, 3])
            {
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }
        }

        private void LoadGridViewDataAndChart(ExcelWorksheet worksheet)
        {
            // GridView Headers
            for (int i = 0; i < gvMonthlyBalance.Columns.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = gvMonthlyBalance.Columns[i].HeaderText;
            }

            // Formatting headers
            using (var range = worksheet.Cells["A1:G1"])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            }

            // GridView Data rows
            for (int i = 0; i < gvMonthlyBalance.Rows.Count; i++)
            {
                for (int j = 0; j < gvMonthlyBalance.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1].Value = gvMonthlyBalance.Rows[i].Cells[j].Text;
                }
            }

            // Formatting data rows
            using (var range = worksheet.Cells[2, 1, gvMonthlyBalance.Rows.Count + 1, gvMonthlyBalance.Columns.Count])
            {
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            // Linear Chart
            var chart = worksheet.Drawings.AddChart("LinearChart", eChartType.LineMarkers);
            chart.Title.Text = "Monthly Balance Linear Chart";
            chart.SetPosition(gvMonthlyBalance.Rows.Count + 3, 0, 0, 0);
            chart.SetSize(800, 400);

            // Add series to chart
            var creditsSeries = chart.Series.Add(
                ExcelRange.GetAddress(2, 2, gvMonthlyBalance.Rows.Count + 1, 2),
                ExcelRange.GetAddress(2, 1, gvMonthlyBalance.Rows.Count + 1, 1));
            creditsSeries.Header = "Total Credits";

            var debitsSeries = chart.Series.Add(
                ExcelRange.GetAddress(2, 3, gvMonthlyBalance.Rows.Count + 1, 3),
                ExcelRange.GetAddress(2, 1, gvMonthlyBalance.Rows.Count + 1, 1));
            debitsSeries.Header = "Total Debits";

            var balanceSeries = chart.Series.Add(
                ExcelRange.GetAddress(2, 6, gvMonthlyBalance.Rows.Count + 1, 6),
                ExcelRange.GetAddress(2, 1, gvMonthlyBalance.Rows.Count + 1, 1));
            balanceSeries.Header = "Remaining Balance";
        }


       
        private void WriteExcelToResponse(ExcelPackage excel, string fileName)
        {
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", $"attachment; filename={fileName}");
            Response.BinaryWrite(excel.GetAsByteArray());
            Response.End();
        }

        private void BindMonthlyBalance(DateTime startDate, DateTime endDate)
        {
            if (ddlAccountNumber.SelectedValue == "0")
            {
                pnlRemBal.Visible = false;
                ErrorMessageLabel.Text = "Please select a valid account number.";
                return;
            }

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(@"spBSA_Savings", con)) 
                {

                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@QueryType", "MonthlySummary");
                    cmd.Parameters.AddWithValue("@AccountID", ddlAccountNumber.SelectedValue);
                    cmd.Parameters.AddWithValue("@StartDate", startDate);
                    cmd.Parameters.AddWithValue("@EndDate", endDate);

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    if (dt.Rows.Count > 0)
                    {
                        gvMonthlyBalance.DataSource = dt;
                        gvMonthlyBalance.DataBind();
                        lblRemBal.Text = "";
                        pnlRemBal.Visible = true;
                    }
                    else
                    {
                        gvMonthlyBalance.DataSource = "";
                        gvMonthlyBalance.DataBind();
                        pnlRemBal.Visible = false;
                        lblRemBal.Text = "No Data Available...";

                    }

                    LoadCandlestickChart(startDate, endDate);
                    LoadTopCredits(startDate, endDate);
                    LoadTopDebits(startDate, endDate);
                    LoanTransactionSummary(startDate, endDate);
                    LoanTransactionDetails(startDate, endDate);
                    MonthlyCredit_DebitTransactions(startDate, endDate);
                    ChequeReturnTransax(startDate, endDate);
                }
            }
        }

        private void LoadCandlestickChart(DateTime startDate, DateTime endDate)
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(@"spBSA_Savings", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@QueryType", "MonthlyCreditSummary");
                    cmd.Parameters.AddWithValue("@AccountID", ddlAccountNumber.SelectedValue);
                    cmd.Parameters.AddWithValue("@StartDate", startDate);
                    cmd.Parameters.AddWithValue("@EndDate", endDate);

                    con.Open();
                    SqlDataReader reader = cmd.ExecuteReader();

                    var dates = new List<string>();
                    var chartData = new List<object>();

                    while (reader.Read())
                    {
                        string monthYearStr = reader.GetString(0);
                        DateTime monthYear = DateTime.ParseExact(monthYearStr, "yyyy-MM", null);
                        string formattedMonthYear = monthYear.ToString("MMM yyyy");
                        decimal totalCredits = reader.IsDBNull(1) ? 0 : reader.GetDecimal(1);
                        decimal totalDebits = reader.IsDBNull(2) ? 0 : reader.GetDecimal(2);
                        decimal remainingBalance = reader.IsDBNull(3) ? 0 : reader.GetDecimal(3);

                        dates.Add(formattedMonthYear);
                        chartData.Add(new
                        {
                            totalCredits,
                            totalDebits,
                            remainingBalance
                        });
                    }

                    reader.Close();
                    con.Close();

                    string datesJson = JsonConvert.SerializeObject(dates);
                    string chartDataJson = JsonConvert.SerializeObject(chartData);

                    // ClientScript.RegisterStartupScript(this.GetType(), "drawCharts", $"drawCandlestickChart({datesJson}, {chartDataJson}); drawLinearChart({datesJson}, {chartDataJson});", true);

                    ClientScript.RegisterStartupScript(this.GetType(), "drawChart", $"drawLinearChart({datesJson}, {chartDataJson});", true);
                   
                }

            }

        }


        private void LoadTopDebits(DateTime startDate, DateTime endDate)
        {
            if (ddlAccountNumber.SelectedIndex == 0)
            {
                ErrorMessageLabel.Text = "Please select a valid account number.";
                return;
            }
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(@"spBSA_Savings", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@QueryType", "Top10Debits");
                    cmd.Parameters.AddWithValue("@AccountID", ddlAccountNumber.SelectedValue);
                    cmd.Parameters.AddWithValue("@StartDate", startDate);
                    cmd.Parameters.AddWithValue("@EndDate", endDate);

                    con.Open();
                    SqlDataReader reader = cmd.ExecuteReader();
                    DataTable dt = new DataTable();
                    dt.Load(reader);

                    GridViewDebits.DataSource = dt;
                    GridViewDebits.DataBind();

                    reader.Close();
                    con.Close();
                }
            }
        }

        private void LoadTopCredits(DateTime startDate, DateTime endDate)
        {
            if (ddlAccountNumber.SelectedIndex == 0)
            {
                ErrorMessageLabel.Text = "Please select a valid account number.";
                return;
            }
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(@"spBSA_Savings", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@QueryType", "Top10Credits");
                    cmd.Parameters.AddWithValue("@AccountID", ddlAccountNumber.SelectedValue);
                    cmd.Parameters.AddWithValue("@StartDate", startDate);
                    cmd.Parameters.AddWithValue("@EndDate", endDate);

                    con.Open();
                    SqlDataReader reader = cmd.ExecuteReader();
                    DataTable dt = new DataTable();
                    dt.Load(reader);
                    if (dt.Rows.Count > 0)
                    {
                        GridViewCredits.DataSource = dt;
                        GridViewCredits.DataBind();
                        lblHighestTransax.Text = "";
                        pnlHighestTransax.Visible = true;
                    }
                    else
                    {
                        GridViewCredits.DataSource = "";
                        GridViewCredits.DataBind();
                        pnlHighestTransax.Visible = false;
                        lblHighestTransax.Text = "No Data Available...";
                    }


                    reader.Close();
                    con.Close();
                }
            }
        }

        protected void GridViewDebits_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                DateTime date = Convert.ToDateTime(DataBinder.Eval(e.Row.DataItem, "value_date"));
                e.Row.Cells[0].Text = date.ToString("dd MMM yyyy");
            }
        }

        protected void GridViewCredits_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                DateTime date = Convert.ToDateTime(DataBinder.Eval(e.Row.DataItem, "value_date"));
                e.Row.Cells[0].Text = date.ToString("dd MMM yyyy");
            }
        }


        private void LoanTransactionSummary(DateTime startDate, DateTime endDate)
        {
            if (ddlAccountNumber.SelectedIndex == 0)
            {
                ErrorMessageLabel.Text = "Please select a valid account number.";
                return;
            }
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand("spBSAPortal", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@CmdType", "loan_transaction_summary");
                    cmd.Parameters.AddWithValue("@AccountID", ddlAccountNumber.SelectedValue);
                    cmd.Parameters.AddWithValue("@StartDate", startDate);
                    cmd.Parameters.AddWithValue("@EndDate", endDate);

                    con.Open();
                    SqlDataReader reader = cmd.ExecuteReader();
                    DataTable dt = new DataTable();
                    dt.Load(reader);
                    if (dt.Rows.Count > 0)
                    {
                        gvLoanSummary.DataSource = dt;
                        gvLoanSummary.DataBind();
                        lblLoanSheet.Text = "";
                        pnlLoanSheet.Visible = true;
                    }
                    else
                    {
                        pnlLoanSheet.Visible = false;
                        gvLoanSummary.DataSource = "";
                        gvLoanSummary.DataBind();
                        lblLoanSheet.Text = "No Data Available...";
                    }

                    reader.Close();
                    con.Close();
                }
            }
        }
        private void LoanTransactionDetails(DateTime startDate, DateTime endDate)
        {
            if (ddlAccountNumber.SelectedIndex == 0)
            {
                ErrorMessageLabel.Text = "Please select a valid account number.";
                return;
            }
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand("spBSAPortal", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@CmdType", "loan_transaction_details");
                    cmd.Parameters.AddWithValue("@AccountID", ddlAccountNumber.SelectedValue);
                    cmd.Parameters.AddWithValue("@StartDate", startDate);
                    cmd.Parameters.AddWithValue("@EndDate", endDate);

                    con.Open();
                    SqlDataReader reader = cmd.ExecuteReader();
                    DataTable dt = new DataTable();
                    dt.Load(reader);

                    gvLoanDetails.DataSource = dt;
                    gvLoanDetails.DataBind();

                    reader.Close();
                    con.Close();
                }
            }
        }


        private void MonthlyCredit_DebitTransactions(DateTime startDate, DateTime endDate)
        {
            DataTable dtDebits = new DataTable();
            DataTable dtCredits = new DataTable();

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                // Query for Summed Debits
                using (SqlCommand cmdDebits = new SqlCommand(@"spBSA_Savings", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@QueryType", "MonthlyDebitSummary");
                    cmdDebits.Parameters.AddWithValue("@AccountID", ddlAccountNumber.SelectedValue);
                    cmdDebits.Parameters.AddWithValue("@StartDate", startDate);
                    cmdDebits.Parameters.AddWithValue("@EndDate", endDate);

                    con.Open();
                    SqlDataAdapter daDebits = new SqlDataAdapter(cmdDebits);
                    daDebits.Fill(dtDebits);
                    gvMonthlyDebits.DataSource = dtDebits;
                    gvMonthlyDebits.DataBind();
                }

                // Query for Summed Credits
                using (SqlCommand cmdCredits = new SqlCommand(@"spBSA_Savings", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@QueryType", "MonthlyCreditSummaryByMonth");
                    cmdCredits.Parameters.AddWithValue("@AccountID", ddlAccountNumber.SelectedValue);
                    cmdCredits.Parameters.AddWithValue("@StartDate", startDate);
                    cmdCredits.Parameters.AddWithValue("@EndDate", endDate);

                    SqlDataAdapter daCredits = new SqlDataAdapter(cmdCredits);
                    daCredits.Fill(dtCredits);
                    gvMonthlyCredits.DataSource = dtCredits;
                    gvMonthlyCredits.DataBind();
                }
            }
        }

        private void ChequeReturnTransax(DateTime startDate, DateTime endDate)
        {
            if (ddlAccountNumber.SelectedValue == "0")
            {
                ErrorMessageLabel.Text = "Please select a valid account number.";
                return;
            }
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(@"spBSA_Savings", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@QueryType", "ReturnTransactions");
                    cmd.Parameters.AddWithValue("@StartDate", startDate);
                    cmd.Parameters.AddWithValue("@EndDate", endDate);

                    con.Open();
                    SqlDataReader reader = cmd.ExecuteReader();
                    DataTable dt = new DataTable();
                    dt.Load(reader);
                    if (dt.Rows.Count > 0)
                    {
                        gvChequeRtnTransax.DataSource = dt;
                        gvChequeRtnTransax.DataBind();
                        lblChequeRtnTransax.Text = "";
                        //pnlHighestTransax.Visible = true;
                    }
                    else
                    {
                        gvChequeRtnTransax.DataSource = "";
                        gvChequeRtnTransax.DataBind();
                        //pnlHighestTransax.Visible = false;
                        lblChequeRtnTransax.Text = "No Data Available...";
                    }


                    reader.Close();
                    con.Close();
                }
            }
        }


        //private void BindExceptionalTransactions()
        //{
        //    if (ddlAccountNumber.SelectedIndex == 0)
        //    {
        //        ErrorMessageLabel.Text = "Please select a valid account number.";
        //        return;
        //    }
        //    DataTable dt = new DataTable();


        //    using (SqlConnection con = new SqlConnection(connectionString))
        //    {
        //        using (SqlCommand cmd = new SqlCommand(@"spBSA_Savings", con))
        //        {
        //            cmd.CommandType = CommandType.StoredProcedure;
        //            cmd.Parameters.AddWithValue("@QueryType", "HighValueTransactions");
        //            con.Open();
        //            SqlDataAdapter da = new SqlDataAdapter(cmd);
        //            da.Fill(dt);
        //            gvExceptionalTransactions.DataSource = dt;
        //            gvExceptionalTransactions.DataBind();
        //        }
        //    }
        //}



    }
}