SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROCEDURE spBSA_Savings
    @AccountId INT = 0,
    @StartDate DATE,
    @EndDate DATE,
    @QueryType VARCHAR(50) = ''
AS
BEGIN
    
    IF @QueryType = 'AccountDetails'
    BEGIN
        -- First Query
        SELECT 
            account_name, 
            account_number, 
            bank_name, 
            account_type 
        FROM 
            bank_account_details 
        WHERE 
            bad_id = @AccountId;
    END
    ELSE IF @QueryType = 'TransactionData'
    BEGIN
        -- Second Query with CTE TransactionData
        WITH TransactionData AS (
            SELECT 
                MIN(value_date) AS StartDate, 
                MAX(value_date) AS EndDate, 
                SUM(credits) AS TotalCredits, 
                MAX(current_balance) AS ClosingBalance
            FROM 
                bank_account_transactions
            WHERE 
                bank_account_id = @AccountId 
                AND value_date BETWEEN @StartDate AND @EndDate
        )
        SELECT 
            StartDate,
            EndDate,
            DATEDIFF(MONTH, StartDate, EndDate) AS NumberOfMonths,
            TotalCredits,
            ClosingBalance
        FROM 
            TransactionData;
    END
    ELSE IF @QueryType = 'MonthlyIncomeExpense'
    BEGIN
        -- Third Query
        SELECT 
            FORMAT(value_date, 'yyyy-MM') AS Month, 
            SUM(credits) AS MonthlyIncome, 
            SUM(debits) AS MonthlyExpense
        FROM 
            bank_account_transactions 
        WHERE  
            (L3_Labels LIKE '%expense%' OR L3_Labels LIKE '%income%') 
            AND bank_account_id = @AccountId 
            AND value_date >= @StartDate 
            AND value_date <= @EndDate
        GROUP BY 
            FORMAT(value_date, 'yyyy-MM')
        ORDER BY 
            FORMAT(value_date, 'yyyy-MM');
    END
    ELSE IF @QueryType = 'TotalIncomeExpense'
    BEGIN
        -- Fourth Query
        SELECT 
            SUM(credits) AS TotalIncome, 
            SUM(debits) AS TotalExpense, 
            (SUM(credits) - SUM(debits)) AS TotalBalance
        FROM 
            bank_account_transactions 
        WHERE  
            (L3_Labels LIKE '%expense%' OR L3_Labels LIKE '%income%') 
            AND bank_account_id = @AccountId 
            AND value_date >= @StartDate 
            AND value_date <= @EndDate;
    END
    ELSE IF @QueryType = 'TransactionDetails'
    BEGIN
        -- Fifth Query
        SELECT 
            value_date, 
            credits, 
            debits 
        FROM 
            bank_account_transactions 
        WHERE  
            (L3_Labels LIKE '%expense%' OR L3_Labels LIKE '%income%') 
            AND value_date >= @StartDate 
            AND value_date <= @EndDate 
            AND bank_account_id = @AccountId;
    END
    ELSE IF @QueryType = 'AllTransactions'
    BEGIN
        -- Sixth Query
        SELECT 
            * 
        FROM 
            bank_account_transactions 
        WHERE 
            bank_account_id = @AccountId 
            AND value_date >= @StartDate 
            AND value_date <= @EndDate;
    END
    ELSE IF @QueryType = 'MonthlySummary'
    BEGIN
        -- Seventh Query
        SELECT
            CONCAT(DATENAME(MONTH, value_date), ' ', YEAR(value_date)) AS MonthYear,
            ISNULL(SUM(BAT.credits), 0) AS TotalCredits,
            ISNULL(SUM(BAT.debits), 0) AS TotalDebits,
            MAX(BAT.current_balance) AS MaxBalance,
            MIN(BAT.current_balance) AS MinBalance,
            SUM(BAT.credits) - SUM(BAT.debits) AS RemainingBalance,
            MIN(BAT.value_date) AS StartDate,
            MAX(BAT.value_date) AS EndDate
        FROM
            bank_account_transactions BAT
        WHERE
            BAT.bank_account_id = @AccountID
            AND BAT.value_date BETWEEN @StartDate AND @EndDate
        GROUP BY
            DATENAME(MONTH, value_date),
            YEAR(value_date),
            MONTH(value_date)
        ORDER BY
            YEAR(value_date),
            MONTH(value_date);
    END
    ELSE IF @QueryType = 'MonthlyCreditSummary'
    BEGIN
        -- Eighth Query
        SELECT
            CONVERT(VARCHAR(7), BAT.value_date, 120) AS MonthYear,
            ISNULL(SUM(BAT.credits), 0) AS TotalCredits,
            ISNULL(SUM(BAT.debits), 0) AS TotalDebits,
            ISNULL(SUM(BAT.credits), 0) - ISNULL(SUM(BAT.debits), 0) AS RemainingBalance
        FROM
            bank_account_transactions BAT
        WHERE
            BAT.bank_account_id = @AccountID
            AND BAT.value_date BETWEEN @StartDate AND @EndDate
        GROUP BY
            CONVERT(VARCHAR(7), BAT.value_date, 120)
        ORDER BY
            MonthYear;
    END
    ELSE IF @QueryType = 'Top10Debits'
    BEGIN
        -- Ninth Query
        SELECT TOP 10
            description, 
            L3_Labels,
            value_date,
            debits AS amount
        FROM
            bank_account_transactions
        WHERE
            bank_account_id = @AccountID
            AND debits > 0
            AND value_date BETWEEN @StartDate AND @EndDate
        ORDER BY
            debits DESC, value_date DESC;
    END
    ELSE IF @QueryType = 'Top10Credits'
    BEGIN
        -- Tenth Query
        SELECT TOP 10
            description, 
            L3_Labels,
            value_date,
            credits AS amount
        FROM
            bank_account_transactions
        WHERE
            bank_account_id = @AccountID
            AND credits > 0 
            AND value_date BETWEEN @StartDate AND @EndDate
        ORDER BY
            credits DESC, value_date DESC;
    END
    ELSE IF @QueryType = 'MonthlyDebitSummary'
    BEGIN
        -- Eleventh Query
        SELECT
            CONCAT(DATENAME(MONTH, value_date), '-', YEAR(value_date)) AS TransactionDate,
            SUM(debits) AS TotalDebits
        FROM
            bank_account_transactions
        WHERE
            debits > 0 
            AND bank_account_id = @AccountID
            AND value_date BETWEEN @StartDate AND @EndDate
        GROUP BY
            DATENAME(MONTH, value_date),
            YEAR(value_date),
            MONTH(value_date)
        ORDER BY
            TotalDebits DESC;
    END
    ELSE IF @QueryType = 'MonthlyCreditSummaryByMonth'
    BEGIN
        -- Twelfth Query
        SELECT
            CONCAT(DATENAME(MONTH, value_date), '-', YEAR(value_date)) AS TransactionDate,
            SUM(credits) AS TotalCredits
        FROM
            bank_account_transactions
        WHERE
            credits > 0 
            AND bank_account_id = @AccountID
            AND value_date BETWEEN @StartDate AND @EndDate
        GROUP BY
            DATENAME(MONTH, value_date),
            YEAR(value_date),
            MONTH(value_date)
        ORDER BY
            TotalCredits DESC;
    END
    ELSE IF @QueryType = 'ReturnTransactions'
    BEGIN
        -- Thirteenth Query
        SELECT 
            value_date,
            L3_Labels,
            description, 
            debits, 
            'Debits' AS TransactionType
        FROM 
            bank_account_transactions 
        WHERE 
            L3_Labels LIKE '%return%' 
            AND value_date BETWEEN @StartDate AND @EndDate;
    END
    ELSE IF @QueryType = 'HighValueTransactions'
    BEGIN
        -- Fourteenth Query
        SELECT 
            'Cash Deposit >=50000' AS Category, 
            COUNT(*) AS Number, 
            'Debit' AS Debit, 
            '' AS Credit
        FROM 
            bank_account_transactions 
        WHERE 
            L3_Labels = 'cash Deposit' 
            AND debits >= 50000
        UNION ALL
        SELECT 
            'Cash Withdrawal >=50000', 
            COUNT(*), 
            '', 
            'Credit'
        FROM 
            bank_account_transactions 
        WHERE 
            L3_Labels = 'cash withdrawal' 
            AND credits >= 50000
        UNION ALL
        SELECT 
            'Cash Deposit >=50% of Average Income', 
            COUNT(*), 
            'Debit', 
            ''
        FROM 
            bank_account_transactions 
        WHERE 
            L3_Labels = 'cash Deposit' 
            AND debits >= 0.5 * (SELECT AVG(debits) FROM bank_account_transactions)
        UNION ALL
        SELECT 
            'Cash Withdrawal >=50% of Average Income', 
            COUNT(*), 
            '', 
            'Credit'
        FROM 
            bank_account_transactions 
        WHERE 
            L3_Labels = 'cash withdrawal' 
            AND credits >= 0.5 * (SELECT AVG(credits) FROM bank_account_transactions)
        UNION ALL
        SELECT 
            'Cash Deposit > Average Income', 
            COUNT(*), 
            'Debit', 
            ''
        FROM 
            bank_account_transactions 
        WHERE 
            L3_Labels = 'cash deposit' 
            AND debits > (SELECT AVG(debits) FROM bank_account_transactions)
        UNION ALL
        SELECT 
            'High Value Transactions', 
            COUNT(*), 
            '', 
            'Credit'
        FROM 
            bank_account_transactions 
        WHERE 
            debits >= 100000 
            OR credits >= 100000
        UNION ALL
        SELECT 
            'ATM Withdrawal not in multiples of 100', 
            COUNT(*), 
            'Debit', 
            ''
        FROM 
            bank_account_transactions 
        WHERE 
            L1_Labels = 'atm' 
            AND debits % 100 <> 0
        UNION ALL
        SELECT 
            'RTGS Payments less than 2 lakhs', 
            COUNT(*), 
            '', 
            'Credit'
        FROM 
            bank_account_transactions 
        WHERE 
            L1_Labels = 'atm' 
            AND credits< 200000
        UNION ALL
        SELECT 
            'Transactions on Sundays', 
            COUNT(*), 
            '', 
            'Credit'
        FROM 
            bank_account_transactions 
        WHERE 
            DATENAME(dw, value_date) = 'Sunday';
    END
END
