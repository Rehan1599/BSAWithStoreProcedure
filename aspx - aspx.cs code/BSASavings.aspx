<%@ Page Title="" Language="C#" MasterPageFile="~/Main.Master" AutoEventWireup="true" CodeBehind="BSASavings.aspx.cs" Inherits="FinAnalyz_NetNew.BSA2" %>

<%@ Register Assembly="System.Web" Namespace="System.Web.UI.WebControls" TagPrefix="asp" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/jquery-datetimepicker/2.5.20/jquery.datetimepicker.min.css" />
    <link rel="stylesheet" href="Libraries/CSS/BSA.css" />
    <%--  <style type="text/css">
        .accordion-button {
            text-align: center; /* Center-align text horizontally */
        }

        .button-text {
            display: block; /* Ensure text occupies full width of button */
            width: 100%; /* Ensure text occupies full width of button */
        }
    </style>--%>

    <style type="text/css">
        .card {
            border: 1px solid #ccc;
            border-radius: 5px;
            padding-top: 0;
            margin-bottom: 20px;
            max-height: 100vh;
            overflow-y: auto;
            width: auto;
        }

        .table th {
            background-color: #007bff;
            color: white;
            position: sticky;
            top: 0; /* Adjust to height of card header */
            z-index: 5;
        }
        
    </style>
     <style>
        /* Loader Text CSS */
        .loader-text {
            position: absolute;
            top: 46%; /* Positioned above the progress bar */
            left: 60%;
            transform: translate(-50%, -50%);
            font-size: 18px;
            color: black;
            text-align: center;
            white-space: nowrap;
        }

        /* Progress Bar CSS */
        .progress-container {
            position: absolute;
            top: 50%; /* Positioned below the text */
            left: 60%;
            transform: translateX(-50%);
            width: 80%; /* Adjust the width as needed */
            max-width: 300px; /* Set a max-width for responsiveness */
            height: 10px;
            background-color: lightgray;
            border-radius: 10px;
            overflow: hidden; /* Ensure the progress bar stays within its container */
        }

        .progress-bar {
            height: 100%;
            background-color: green;
            border-radius: 10px;
            text-align: center;
            line-height: 20px; /* Adjust according to progress bar height */
            color: white; /* Text color */
            transition: width 0.3s ease; /* Smooth transition for width changes */
        }

        /* Hide content initially */
        #content {
            display: none;
            margin-top: 20px; /* Space below progress bar */
           
        }
    </style>


    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">

 
      <div class="loader-text" id="loader-text"></div>
    <div class="progress-container" id="progress-container">
        <div class="progress">
            <div class="progress-bar" id="progress-bar" role="progressbar" style="width: 0%;" aria-valuenow="0" aria-valuemin="0" aria-valuemax="100">0%</div>
        </div>
    </div>
    
 
    <div id="content" >
        <h1 class="text-center">BSA SAVING ACCOUNTS</h1>
        <div class="row">

            <div class="col-md-4">
                <asp:DropDownList ID="ddlAccountNumber" runat="server" OnSelectedIndexChanged="ddlAccountNumber_SelectedIndexChanged" AutoPostBack="true" CssClass="form-control">
                </asp:DropDownList>
            </div>

            <div class="col-md-4">
                <div class="input-group mb-3 mx-auto">
                    <asp:TextBox CssClass="form-control" Enabled="false" placeholder="Enter Start Date" ID="txtStartDate" runat="server"></asp:TextBox>
                    <asp:TextBox CssClass="form-control" Enabled="false" placeholder="Enter End Date" ID="txtEndDate" runat="server"></asp:TextBox>
                    <div class="input-group-append">
                        <asp:LinkButton ID="SearchBtn" Enabled="false" CssClass="btn btn-outline-success" OnClick="SearchBtn_Click" runat="server">
                        <i class="fas fa-search"></i>
                        </asp:LinkButton>
                    </div>
                </div>
            </div>
            <div class="col-md-4">
                <%--            <div class="btn-group btn-group-justified d-flex">
               <asp:Button ID="btnLastFinancialYear" CssClass="btn btn-outline-dark d-block" runat="server" Text="Last Financial Year" OnClick="btnLastFinancialYear_Click" />
           </div>--%>


                <asp:DropDownList ID="ddlTimeFrame" Enabled="false" runat="server" AutoPostBack="true" CssClass="form-control" OnSelectedIndexChanged="ddlTimeFrame_SelectedIndexChanged">
                    <asp:ListItem Value="0" Text="-- Select Date Range --" />
                    <asp:ListItem Value="today" Text="Today" />
                    <asp:ListItem Value="yesterday" Text="Yesterday" />
                    <asp:ListItem Value="thisWeek" Text="This Week" />
                    <asp:ListItem Value="lastWeek" Text="Last Week" />
                    <asp:ListItem Value="thisMonth" Text="This Month" />
                    <asp:ListItem Value="lastMonth" Text="Last Month" />
                    <asp:ListItem Value="last3Months" Text="Last 3 Months" />
                    <asp:ListItem Value="last6Months" Text="Last 6 Months" />
                    <asp:ListItem Value="last12Months" Text="Last 12 Months" />
                    <asp:ListItem Value="lastFinancialYear" Text="Last Financial Year" />
                </asp:DropDownList>
            </div>

        </div>
        <h6>
            <asp:Label ID="ErrorMessageLabel" runat="server" CssClass="text-danger"></asp:Label>
        </h6>

        <div>
            <asp:Button Visible="true" ID="btnDownloadExcel" runat="server" CssClass="btn btn-primary" Text="Download Excel" OnClick="btnDownloadExcel_Click" />
        </div>

        <div class="accordion" id="accordionPanelsStayOpenExample">
            <div class="accordion-item">
                <h2 class="accordion-header">
                    <button class="accordion-button" type="button" data-bs-toggle="collapse" data-bs-target="#panelsStayOpen-collapseOne" aria-expanded="true" aria-controls="panelsStayOpen-collapseOne">
                        <h5 class="button-text">Bank Statement Analysis Report</h5>
                    </button>
                </h2>
                <div id="panelsStayOpen-collapseOne" class="accordion-collapse collapse show">
                    <div class="accordion-body">
                        <h5 class="text-center" id="panelHide" runat="server">Please Select Account Number First...</h5>
                        <asp:Panel ID="BankDetailsPanel" runat="server" Visible="false" CssClass="bank-details-panel">
                            <%--<asp:Button ID="btnExportBankDetails" CssClass="btn btn-primary" runat="server" Text="Export Bank Details to Excel" OnClick="btnExportBankDetails_Click" />--%>
                            <table class="table table-striped table-borderless table-hover">
                                <tr>
                                    <td class="table-header">Account Name</td>
                                    <td>
                                        <asp:Label ID="lblAccountName" runat="server"></asp:Label></td>
                                </tr>
                                <tr>
                                    <td class="table-header">Account No.</td>
                                    <td>
                                        <asp:Label ID="lblAccountNumber" runat="server"></asp:Label></td>
                                </tr>
                                <tr>
                                    <td class="table-header">Bank Name</td>
                                    <td>
                                        <asp:Label ID="lblBankName" runat="server"></asp:Label></td>
                                </tr>

                                <tr>
                                    <td class="table-header">Period</td>
                                    <td>
                                        <asp:Label ID="lblPeriod" runat="server"></asp:Label></td>
                                </tr>
                                <tr>
                                    <td class="table-header">No. of Months</td>
                                    <td>
                                        <asp:Label ID="lblNumberOfMonths" runat="server"></asp:Label></td>
                                </tr>
                                <tr>
                                    <td class="table-header">Account Type</td>
                                    <td>
                                        <asp:Label ID="lblAccountType" runat="server"></asp:Label></td>
                                </tr>
                                <tr>
                                    <td class="table-header">Currency</td>
                                    <td>
                                        <asp:Label ID="lblCurrency" runat="server"></asp:Label></td>
                                </tr>
                                <tr>
                                    <td class="table-header">Opening Balance</td>
                                    <td>
                                        <asp:Label ID="lblOpeningBalance" runat="server"></asp:Label></td>
                                </tr>
                                <tr>
                                    <td class="table-header">Closing Balance</td>
                                    <td>
                                        <asp:Label ID="lblClosingBalance" runat="server"></asp:Label></td>
                                </tr>
                            </table>
                        </asp:Panel>
                    </div>
                </div>
            </div>

            <div class="accordion-item">
                <h2 class="accordion-header">
                    <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#panelsStayOpen-collapseTwo" aria-expanded="false" aria-controls="panelsStayOpen-collapseTwo">
                        <h5 class="button-text">Transaction Details</h5>
                    </button>
                </h2>
                <div id="panelsStayOpen-collapseTwo" class="accordion-collapse collapse">
                    <div class="accordion-body">

                        <div class="card">
                            <div class="card-body" style="max-height: 70vh">
                                <h5 class="text-center" id="hidetxt" runat="server">Please Select Account Number First...</h5>

                                <h4 class="text-center">
                                    <asp:Label ID="lblTransactionDetails" runat="server"></asp:Label></h4>
                                <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"
                                    CssClass="table table-striped table-hover">
                                    <Columns>
                                        <asp:TemplateField HeaderText="ID" SortExpression="id">
                                            <ItemTemplate>
                                                <asp:Label ID="lblID" runat="server" Text='<%# Eval("id") %>'></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateField>

                                        <asp:TemplateField HeaderText="Transaction Type" SortExpression="transaction_type">
                                            <ItemTemplate>
                                                <asp:Label ID="lblTransactionType" runat="server" Text='<%# Eval("transaction_type") %>'></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateField>

                                        <asp:TemplateField HeaderText="Debits" SortExpression="debits">
                                            <ItemTemplate>
                                                <asp:Label ID="lblDebits" runat="server" Text='<%# Eval("debits", "{0:C}") %>'></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateField>

                                        <asp:TemplateField HeaderText="Credits" SortExpression="credits">
                                            <ItemTemplate>
                                                <asp:Label ID="lblCredits" runat="server" Text='<%# Eval("credits", "{0:C}") %>'></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateField>

                                        <asp:TemplateField HeaderText="Current Balance" SortExpression="current_balance">
                                            <ItemTemplate>
                                                <asp:Label ID="lblCurrentBalance" runat="server" Text='<%# Eval("current_balance", "{0:C}") %>'></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateField>

                                        <asp:TemplateField HeaderText="Transaction Time" SortExpression="transaction_time">
                                            <ItemTemplate>
                                                <asp:Label ID="lblTransactionTime" runat="server" Text='<%# Eval("transaction_time") %>'></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateField>

                                        <asp:TemplateField HeaderText="Date" SortExpression="value_date">
                                            <ItemTemplate>
                                                <asp:Label ID="lblDate" runat="server" Text='<%# Eval("value_date", "{0:dd-M-yyyy}") %>'></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                    </Columns>
                                </asp:GridView>

                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <div class="accordion-item">
                <h2 class="accordion-header">
                    <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#panelsStayOpen-collapseThree" aria-expanded="false" aria-controls="panelsStayOpen-collapseThree">
                        <h5 class="button-text">EOD Analisys</h5>
                    </button>
                </h2>
                <div id="panelsStayOpen-collapseThree" class="accordion-collapse collapse">
                    <div class="accordion-body">
                        <h5 class="text-center" id="txtEOD" runat="server">Please Select Account Number First...</h5>
                        <h5 class="text-center">
                            <asp:Label ID="lblEOD" runat="server"></asp:Label></h5>
                        <asp:Panel ID="EODPanel" runat="server" Visible="false">
                            <div id="chartDiv" style="width: 100%; height: 600px;"></div>
                            <br />
                            <h1 class="text-center">Monthly Income And Expense</h1>
                            <div>
                                <asp:GridView ID="MonthlySummaryGridView" OnRowDataBound="MonthlySummaryGridView_RowDataBound" runat="server" AutoGenerateColumns="False" CssClass="table table-striped table-hover">
                                    <Columns>
                                        <asp:BoundField DataField="Month" HeaderText="Month" DataFormatString="{0:MMMM yyyy}" />
                                        <asp:BoundField DataField="MonthlyIncome" HeaderText="Monthly Income" DataFormatString="{0:C}" />
                                        <asp:BoundField DataField="MonthlyExpense" HeaderText="Monthly Expense" DataFormatString="{0:C}" />
                                    </Columns>

                                </asp:GridView>

                            </div>
                            <hr />
                            <div>
                                <h1 class="text-center">Total Remaining Balance</h1>
                                <asp:GridView ID="SummaryGridView" runat="server" AutoGenerateColumns="False" CssClass="table table-striped table-hover">
                                    <Columns>
                                        <asp:BoundField DataField="TotalIncome" HeaderText="Total Income" DataFormatString="{0:C}" />
                                        <asp:BoundField DataField="TotalExpense" HeaderText="Total Expense" DataFormatString="{0:C}" />
                                        <asp:BoundField DataField="TotalBalance" HeaderText="Remaining Balance" DataFormatString="{0:C}" />
                                    </Columns>
                                </asp:GridView>

                            </div>

                        </asp:Panel>
                    </div>
                </div>
            </div>

            <div class="accordion-item">
                <h2 class="accordion-header">
                    <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#panelsStayOpen-collapseFour" aria-expanded="false" aria-controls="panelsStayOpen-collapseFour">
                        <h5>Monthly Remaining Balance</h5>
                    </button>
                </h2>
                <div id="panelsStayOpen-collapseFour" class="accordion-collapse collapse">
                    <div class="accordion-body">
                        <h5 class="text-center" id="txtRemBal" runat="server">Please Select Account Number First...</h5>
                        <h5 class="text-center">
                            <asp:Label ID="lblRemBal" runat="server"></asp:Label></h5>
                        <asp:Panel runat="server" ID="pnlRemBal" Visible="false">
                            <div id="linearChartContainer" style="width: 100%; height: 400px; margin-top: 20px;"></div>
                            <br />
                          
                            <asp:GridView ID="gvMonthlyBalance" runat="server" AutoGenerateColumns="False" CssClass="table table-striped table-hover">
                                <Columns>
                                    <asp:BoundField DataField="MonthYear" HeaderText="Month Year" DataFormatString="{0:MMM yyyy}" />
                                    <asp:BoundField DataField="TotalCredits" HeaderText="Total Credits" DataFormatString="{0:C}" />
                                    <asp:BoundField DataField="TotalDebits" HeaderText="Total Debits" DataFormatString="{0:C}" />
                                    <asp:BoundField DataField="MaxBalance" HeaderText="Max Balance" DataFormatString="{0:C}" />
                                    <asp:BoundField DataField="MinBalance" HeaderText="Min Balance" DataFormatString="{0:C}" />
                                    <asp:BoundField DataField="RemainingBalance" HeaderText="Remaining Balance" DataFormatString="{0:C}" />
                                </Columns>
                            </asp:GridView>
                        </asp:Panel>


                    </div>
                </div>
            </div>

            <div class="accordion-item">
                <h2 class="accordion-header">
                    <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#panelsStayOpen-collapseFive" aria-expanded="false" aria-controls="panelsStayOpen-collapseFive">
                        <h5>10 Highest Credit & Debit Transactions</h5>
                    </button>
                </h2>
                <div id="panelsStayOpen-collapseFive" class="accordion-collapse collapse">
                    <div class="accordion-body">
                        <h5 class="text-center" id="txtHighestTransax" runat="server">Please Select Account Number First...</h5>
                        <h5 class="text-center">
                            <asp:Label ID="lblHighestTransax" runat="server"></asp:Label></h5>

                        <asp:Panel ID="pnlHighestTransax" runat="server" Visible="false">



                            <h3 class="text-center">10 Highest Credit Transactions</h3>
                            <asp:GridView ID="GridViewCredits" runat="server" AutoGenerateColumns="False" CssClass="table table-striped table-hover" OnRowDataBound="GridViewCredits_RowDataBound">
                                <Columns>
                                    <asp:BoundField DataField="value_date" HeaderText="Transaction Date" ItemStyle-Width="10%" />
                                    <asp:BoundField DataField="description" HeaderText="Description" ItemStyle-Width="50%" />
                                    <asp:BoundField DataField="L3_Labels" HeaderText="Transaction Type" ItemStyle-Width="20%" />
                                    <asp:BoundField DataField="amount" HeaderText="Credit" DataFormatString="{0:C}" ItemStyle-Width="20%" />
                                </Columns>
                            </asp:GridView>

                            <br />
                            <hr />
                            <br />

                            <h3 class="text-center">10 Highest Debit Transactions</h3>
                            <asp:GridView ID="GridViewDebits" runat="server" AutoGenerateColumns="False" OnRowDataBound="GridViewDebits_RowDataBound" CssClass="table table-striped table-hover">
                                <Columns>
                                    <asp:BoundField DataField="value_date" HeaderText="Transaction Date" ItemStyle-Width="10%" />
                                    <asp:BoundField DataField="description" HeaderText="Description" ItemStyle-Width="50%" />
                                    <asp:BoundField DataField="L3_Labels" HeaderText="Transaction Type" ItemStyle-Width="20%" />
                                    <asp:BoundField DataField="amount" HeaderText="Debit" DataFormatString="{0:C}" ItemStyle-Width="20%" />
                                </Columns>
                            </asp:GridView>
                            <br />
                            <hr />
                            <br />
                            <div class="row">
                                <div class="col-6">
                                    <h2 class="text-center">Monthly Debits Transactions</h2>
                                    <asp:GridView ID="gvMonthlyDebits" runat="server" AutoGenerateColumns="false" CssClass="table table-striped table-hover">
                                        <Columns>
                                            <asp:BoundField DataField="TransactionDate" HeaderText="Month-Year" />
                                            <asp:BoundField DataField="TotalDebits" HeaderText="Total Debits" DataFormatString="{0:C}" />
                                        </Columns>
                                    </asp:GridView>
                                </div>

                                <div class="col-6">
                                    <h2 class="text-center">Monthly Credits Transactions</h2>
                                    <asp:GridView ID="gvMonthlyCredits" runat="server" AutoGenerateColumns="false" CssClass="table table-striped table-hover">
                                        <Columns>
                                            <asp:BoundField DataField="TransactionDate" HeaderText="Month-Year" />
                                            <asp:BoundField DataField="TotalCredits" HeaderText="Total Credits" DataFormatString="{0:C}" />
                                        </Columns>
                                    </asp:GridView>
                                </div>

                            </div>

                        </asp:Panel>

                    </div>
                </div>
            </div>

            <div class="accordion-item">
                <h2 class="accordion-header">
                    <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#panelsStayOpen-collapseSix" aria-expanded="false" aria-controls="panelsStayOpen-collapseSix">
                        <h5>Loan Sheet</h5>
                    </button>
                </h2>
                <div id="panelsStayOpen-collapseSix" class="accordion-collapse collapse">
                    <div class="accordion-body">
                        <h5 class="text-center" id="txtLoanSheet" runat="server">Please Select Account Number First...</h5>
                        <h5 class="text-center">
                            <asp:Label ID="lblLoanSheet" runat="server"></asp:Label></h5>

                        <asp:Panel ID="pnlLoanSheet" runat="server" Visible="false">
                            <h1 class="text-center">Summary of Loan Transactions</h1>
                            <asp:GridView ID="gvLoanSummary" runat="server" AutoGenerateColumns="False" CssClass="table table-striped table-hover">
                                <Columns>
                                    <asp:BoundField DataField="TransactionDate" HeaderText="Month" />
                                    <asp:BoundField DataField="TotalCredits" HeaderText="Credits" DataFormatString="{0:C}" />
                                    <asp:BoundField DataField="TotalDebits" HeaderText="Debits" DataFormatString="{0:C}" />
                                </Columns>
                            </asp:GridView>
                            <br />
                            <hr />
                            <br />
                            <h1 class="text-center">Details of Loan Summary</h1>
                            <asp:GridView ID="gvLoanDetails" runat="server" AutoGenerateColumns="False" CssClass="table table-striped table-hover">
                                <Columns>
                                    <asp:BoundField DataField="SerialNo" HeaderText="Serial No." />
                                    <asp:BoundField DataField="value_date" HeaderText="Date" DataFormatString="{0:dd MMM yyyy}" />
                                    <asp:BoundField DataField="L3_Labels" HeaderText="Transaction Type" />
                                    <asp:BoundField DataField="TotalCredits" HeaderText="Credits" DataFormatString="{0:C}" />
                                    <asp:BoundField DataField="TotalDebits" HeaderText="Debits" DataFormatString="{0:C}" />
                                </Columns>
                            </asp:GridView>
                        </asp:Panel>
                    </div>
                </div>
            </div>

            <div class="accordion-item">
                <h2 class="accordion-header">
                    <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#panelsStayOpen-collapseSeven" aria-expanded="false" aria-controls="panelsStayOpen-collapseSeven">
                        <h5>Cheque Return Transactions</h5>
                    </button>
                </h2>
                <div id="panelsStayOpen-collapseSeven" class="accordion-collapse collapse">
                    <div class="accordion-body">

                        <h5 class="text-center" id="txtChequeRtnTransax" runat="server">Please Select Account Number First...</h5>
                        <h5 class="text-center">
                            <asp:Label ID="lblChequeRtnTransax" runat="server"></asp:Label></h5>

                        <asp:GridView ID="gvChequeRtnTransax" runat="server" AutoGenerateColumns="false" CssClass="table table-striped table-hover">
                            <Columns>
                                <asp:BoundField DataField="value_date" HeaderText="Date" DataFormatString="{0:dd MMM yyyy}" />
                                <asp:BoundField DataField="L3_Labels" HeaderText="Particulars" />
                                <asp:BoundField DataField="description" HeaderText="Reason" />
                                <asp:BoundField DataField="debits" HeaderText="Amount" DataFormatString="{0:C}" />
                                <asp:BoundField DataField="TransactionType" HeaderText="Transaction Type" />
                            </Columns>
                        </asp:GridView>
                    </div>
                </div>
            </div>

            <%-- <div class="accordion-item">
           <h2 class="accordion-header">
               <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#panelsStayOpen-collapseEight" aria-expanded="false" aria-controls="panelsStayOpen-collapseEight">
                   <h5>Summary of Exceptional Transactions</h5>
               </button>
           </h2>
           <div id="panelsStayOpen-collapseEight" class="accordion-collapse collapse">
               <div class="accordion-body">
                   <h1 class="text-center"></h1>
                   <asp:GridView ID="gvExceptionalTransactions" runat="server" AutoGenerateColumns="false" CssClass="table table-striped table-hover">
                       <Columns>
                           <asp:BoundField DataField="Category" HeaderText="Category" />
                           <asp:BoundField DataField="Number" HeaderText="Number" />
                           <asp:BoundField DataField="Debit" HeaderText="Debit" />
                           <asp:BoundField DataField="Credit" HeaderText="Credit" />
                       </Columns>
                   </asp:GridView>
               </div>
           </div>
       </div>--%>
        </div>

    </div>


     <script>
         function typeWriter(text, elementId, delay = 20, callback) {
             let i = 0;
             function type() {
                 if (i < text.length) {
                     document.getElementById(elementId).innerHTML += text.charAt(i);
                     i++;
                     updateProgressBar(i / text.length * 100); // Update progress bar
                     setTimeout(type, delay);
                 } else {
                     if (callback) callback();
                 }
             }
             type();
         }

         function updateProgressBar(percentage) {
             const progressBar = document.getElementById('progress-bar');
             progressBar.style.width = percentage + '%';
             progressBar.setAttribute('aria-valuenow', percentage);
             progressBar.innerHTML = percentage.toFixed(0) + '%'; // Display percentage inside the progress bar
         }

         window.onload = function () {
             const loaderText = "Saving Account Data Loading...";
             const loaderTextElement = document.getElementById('loader-text');
             const progressContainer = document.getElementById('progress-container');

             typeWriter(loaderText, 'loader-text', 20, function () {
                 setTimeout(function () {
                     document.getElementById('loader-text').style.display = 'none';
                     document.querySelector('.progress-container').style.display = 'none';
                     document.getElementById('content').style.display = 'block';
                 }, 500); // Display the loader for at least 500 milliseconds after typing completes
             });

             // Set the width of the progress bar container to match the text width
             setTimeout(function () {
                 progressContainer.style.width = loaderTextElement.offsetWidth + '%';
             }, 0);
         };
    </script>

    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js"></script>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery-datetimepicker/2.5.20/jquery.datetimepicker.full.min.js"></script>
    <script type="text/javascript" src="Library/JS/BSAChart.js"></script>

    <script>
        $(function () {
            $("#<%= txtStartDate.ClientID %>").datetimepicker({
                timepicker: false,
                format: 'd-m-Y',
                scrollMonth: false,
                scrollInput: false,
                yearStart: 2000,
                yearEnd: 2100,
                changeMonth: true,
                changeYear: true,
            });
            $("#<%= txtEndDate.ClientID %>").datetimepicker({
                timepicker: false,
                format: 'd-m-Y',
                scrollMonth: false,
                scrollInput: false,
                yearStart: 2000,
                yearEnd: 2100,
                changeMonth: true,
                changeYear: true,
            });

            // Attach event listener to resize chart when accordion is shown
            document.getElementById('panelsStayOpen-collapseThree').addEventListener('shown.bs.collapse', function () {
                Plotly.Plots.resize('chartDiv');
            });

            document.getElementById('panelsStayOpen-collapseFour').addEventListener('shown.bs.collapse', function () {
                Plotly.Plots.resize('linearChartContainer');
                Plotly.Plots.resize('chartContainer');

            });

            // Also resize the chart when window is resized to handle edge cases
            window.addEventListener('resize', function () {
                Plotly.Plots.resize('chartDiv');
                Plotly.Plots.resize('linearChartContainer');
                Plotly.Plots.resize('chartContainer');
            });


        });
    </script>



</asp:Content>
