/*
SikaLink 2025 performance review
SQL Server data preparation and analysis

Source tables: raw_customers, raw_transactions, raw_agents, raw_loans and
raw_repayments in the sikalink database. Import the source CSV files first.

The raw tables are retained. Clean tables are created only when they do not
already exist; existing clean tables are not overwritten. Run the validation
queries before using the analysis and check any unexpected results.

Transaction measures use successful transactions unless otherwise stated.
Repayment performance includes amounts due by 31 December 2025. A repayment
is on time when it is fully paid by its due date; unpaid and partially paid
amounts remain in the denominator. The default rate uses loan status and
includes groups with at least 30 loans.
*/

USE [sikalink];
GO

/* =============================================================================
   SECTION A: INSPECT THE RAW TABLES
   Compare total rows and distinct business IDs. Do not edit the raw tables.
   ============================================================================= */

SELECT 'raw_customers' AS table_name, COUNT(*) AS total_rows,
       COUNT(DISTINCT customer_id) AS distinct_ids,
       SUM(CASE WHEN customer_id IS NULL OR LTRIM(RTRIM(CONVERT(nvarchar(100), customer_id))) = ''
                THEN 1 ELSE 0 END) AS missing_ids
FROM dbo.raw_customers
UNION ALL
SELECT 'raw_transactions', COUNT(*), COUNT(DISTINCT txn_id),
       SUM(CASE WHEN txn_id IS NULL OR LTRIM(RTRIM(CONVERT(nvarchar(100), txn_id))) = '' THEN 1 ELSE 0 END)
FROM dbo.raw_transactions
UNION ALL
SELECT 'raw_agents', COUNT(*), COUNT(DISTINCT agent_id),
       SUM(CASE WHEN agent_id IS NULL OR LTRIM(RTRIM(CONVERT(nvarchar(100), agent_id))) = '' THEN 1 ELSE 0 END)
FROM dbo.raw_agents
UNION ALL
SELECT 'raw_loans', COUNT(*), COUNT(DISTINCT loan_id),
       SUM(CASE WHEN loan_id IS NULL OR LTRIM(RTRIM(CONVERT(nvarchar(100), loan_id))) = '' THEN 1 ELSE 0 END)
FROM dbo.raw_loans
UNION ALL
SELECT 'raw_repayments', COUNT(*), COUNT(DISTINCT repayment_id),
       SUM(CASE WHEN repayment_id IS NULL OR LTRIM(RTRIM(CONVERT(nvarchar(100), repayment_id))) = '' THEN 1 ELSE 0 END)
FROM dbo.raw_repayments;
GO

-- Examine important blanks without changing them in the raw data.
SELECT 'Customer age group' AS field_name,
       SUM(CASE WHEN age_group IS NULL OR
           LTRIM(RTRIM(CONVERT(nvarchar(100), age_group))) = ''
           THEN 1 ELSE 0 END) AS missing_rows
FROM dbo.raw_customers
UNION ALL
SELECT 'Transaction agent ID',
       SUM(CASE WHEN agent_id IS NULL OR
           LTRIM(RTRIM(CONVERT(nvarchar(100), agent_id))) = ''
           THEN 1 ELSE 0 END)
FROM dbo.raw_transactions
UNION ALL
SELECT 'Repayment paid date',
       SUM(CASE WHEN paid_date IS NULL OR
           LTRIM(RTRIM(CONVERT(nvarchar(100), paid_date, 23))) = ''
           THEN 1 ELSE 0 END)
FROM dbo.raw_repayments;
GO

-- Check text-format issues that need date or currency conversion. If the
-- import wizard already converted a column's type, these raw-table counts may
-- differ from counts in the original CSV; record what you actually observed.
SELECT 'Transaction dates containing a slash' AS issue_found, COUNT(*) AS affected_rows
FROM dbo.raw_transactions
WHERE CHARINDEX('/', CONVERT(varchar(50), txn_date, 23)) > 0
UNION ALL
SELECT 'Transaction amounts containing GHS', COUNT(*)
FROM dbo.raw_transactions
WHERE UPPER(CONVERT(nvarchar(100), amount_ghs)) LIKE '%GHS%';
GO

/* =============================================================================
   SECTION B: CREATE CLEAN TABLES IF MISSING

   SELECT DISTINCT removes exact duplicate records after standardisation; it
   does not decide which conflicting records sharing an ID should be kept.
   If an ID is still duplicated, investigate before using the analysis.
   ============================================================================= */

-- B1. Customers: trim text, standardise region, preserve unknown age groups,
-- and convert dates. Date text with slashes uses DD/MM/YYYY (style 103);
-- ISO date text uses YYYY-MM-DD (style 23).
IF OBJECT_ID(N'dbo.customers') IS NULL
BEGIN
    SELECT DISTINCT
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), customer_id))), '') AS customer_id,
        CASE
            WHEN CHARINDEX('/', CONVERT(varchar(50), signup_date, 23)) > 0
                THEN TRY_CONVERT(date, LEFT(LTRIM(RTRIM(CONVERT(varchar(50), signup_date, 23))), 10), 103)
            ELSE TRY_CONVERT(date, LEFT(LTRIM(RTRIM(CONVERT(varchar(50), signup_date, 23))), 10), 23)
        END AS signup_date,
        UPPER(NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), region))), '')) AS region,
        COALESCE(NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), age_group))), ''), 'Unknown') AS age_group,
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), gender))), '') AS gender,
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), acquisition_channel))), '') AS acquisition_channel,
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), kyc_level))), '') AS kyc_level
    INTO dbo.customers
    FROM dbo.raw_customers;
    PRINT 'Created dbo.customers from dbo.raw_customers.';
END
ELSE
    PRINT 'Skipped dbo.customers: an object with this name already exists.';
GO

-- B2. Transactions: remove exact duplicates, retain missing agents as NULL,
-- parse mixed dates, and remove GHS and comma separators from numeric values.
IF OBJECT_ID(N'dbo.transactions') IS NULL
BEGIN
    SELECT DISTINCT
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), txn_id))), '') AS txn_id,
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), customer_id))), '') AS customer_id,
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), agent_id))), '') AS agent_id,
        CASE
            WHEN CHARINDEX('/', CONVERT(varchar(50), txn_date, 23)) > 0
                THEN TRY_CONVERT(date, LEFT(LTRIM(RTRIM(CONVERT(varchar(50), txn_date, 23))), 10), 103)
            ELSE TRY_CONVERT(date, LEFT(LTRIM(RTRIM(CONVERT(varchar(50), txn_date, 23))), 10), 23)
        END AS txn_date,
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), txn_type))), '') AS txn_type,
        TRY_CONVERT(decimal(19,2),
            REPLACE(REPLACE(UPPER(LTRIM(RTRIM(CONVERT(nvarchar(100), amount_ghs)))), 'GHS', ''), ',', '')) AS amount_ghs,
        TRY_CONVERT(decimal(19,2),
            REPLACE(REPLACE(UPPER(LTRIM(RTRIM(CONVERT(nvarchar(100), fee_ghs)))), 'GHS', ''), ',', '')) AS fee_ghs,
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), channel))), '') AS channel,
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), status))), '') AS status
    INTO dbo.transactions
    FROM dbo.raw_transactions;
    PRINT 'Created dbo.transactions from dbo.raw_transactions.';
END
ELSE
    PRINT 'Skipped dbo.transactions: an object with this name already exists.';
GO

-- B3. Agents: keep an independent, readable copy for analysis and Power BI.
IF OBJECT_ID(N'dbo.agents') IS NULL
BEGIN
    SELECT DISTINCT
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), agent_id))), '') AS agent_id,
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), agent_region))), '') AS agent_region,
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), agent_type))), '') AS agent_type
    INTO dbo.agents
    FROM dbo.raw_agents;
    PRINT 'Created dbo.agents from dbo.raw_agents.';
END
ELSE
    PRINT 'Skipped dbo.agents: an object with this name already exists.';
GO

-- B4. Loans: convert dates, principal, term and interest rate to useful types.
IF OBJECT_ID(N'dbo.loans') IS NULL
BEGIN
    SELECT DISTINCT
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), loan_id))), '') AS loan_id,
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), customer_id))), '') AS customer_id,
        CASE
            WHEN CHARINDEX('/', CONVERT(varchar(50), disbursement_date, 23)) > 0
                THEN TRY_CONVERT(date, LEFT(LTRIM(RTRIM(CONVERT(varchar(50), disbursement_date, 23))), 10), 103)
            ELSE TRY_CONVERT(date, LEFT(LTRIM(RTRIM(CONVERT(varchar(50), disbursement_date, 23))), 10), 23)
        END AS disbursement_date,
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), product))), '') AS product,
        TRY_CONVERT(decimal(19,2),
            REPLACE(REPLACE(UPPER(LTRIM(RTRIM(CONVERT(nvarchar(100), principal_ghs)))), 'GHS', ''), ',', '')) AS principal_ghs,
        TRY_CONVERT(int, term_days) AS term_days,
        TRY_CONVERT(decimal(9,4), interest_rate) AS interest_rate,
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), status))), '') AS status
    INTO dbo.loans
    FROM dbo.raw_loans;
    PRINT 'Created dbo.loans from dbo.raw_loans.';
END
ELSE
    PRINT 'Skipped dbo.loans: an object with this name already exists.';
GO

-- B5. Repayments: keep blank paid_date as NULL; it is not automatically a
-- default. Store amounts as decimal and allow for partial payments.
IF OBJECT_ID(N'dbo.repayments') IS NULL
BEGIN
    SELECT DISTINCT
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), repayment_id))), '') AS repayment_id,
        NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), loan_id))), '') AS loan_id,
        CASE
            WHEN CHARINDEX('/', CONVERT(varchar(50), due_date, 23)) > 0
                THEN TRY_CONVERT(date, LEFT(LTRIM(RTRIM(CONVERT(varchar(50), due_date, 23))), 10), 103)
            ELSE TRY_CONVERT(date, LEFT(LTRIM(RTRIM(CONVERT(varchar(50), due_date, 23))), 10), 23)
        END AS due_date,
        CASE
            WHEN CHARINDEX('/', CONVERT(varchar(50), paid_date, 23)) > 0
                THEN TRY_CONVERT(date, LEFT(LTRIM(RTRIM(CONVERT(varchar(50), paid_date, 23))), 10), 103)
            ELSE TRY_CONVERT(date, LEFT(LTRIM(RTRIM(CONVERT(varchar(50), paid_date, 23))), 10), 23)
        END AS paid_date,
        TRY_CONVERT(decimal(19,2),
            REPLACE(REPLACE(UPPER(LTRIM(RTRIM(CONVERT(nvarchar(100), amount_due_ghs)))), 'GHS', ''), ',', '')) AS amount_due_ghs,
        TRY_CONVERT(decimal(19,2),
            REPLACE(REPLACE(UPPER(LTRIM(RTRIM(CONVERT(nvarchar(100), amount_paid_ghs)))), 'GHS', ''), ',', '')) AS amount_paid_ghs
    INTO dbo.repayments
    FROM dbo.raw_repayments;
    PRINT 'Created dbo.repayments from dbo.raw_repayments.';
END
ELSE
    PRINT 'Skipped dbo.repayments: an object with this name already exists.';
GO

/* =============================================================================
   SECTION C: VERIFY CLEAN TABLES AND RELATIONSHIPS
   Missing dates/numbers can indicate a failed conversion. Inspect the source
   before correcting them. Do not use the analysis until critical issues are
   understood. The raw tables remain available for investigation.
   ============================================================================= */

SELECT 'customers' AS table_name, COUNT(*) AS clean_rows,
       COUNT(DISTINCT customer_id) AS distinct_ids,
       SUM(CASE WHEN customer_id IS NULL THEN 1 ELSE 0 END) AS missing_ids
FROM dbo.customers
UNION ALL
SELECT 'transactions', COUNT(*), COUNT(DISTINCT txn_id),
       SUM(CASE WHEN txn_id IS NULL THEN 1 ELSE 0 END)
FROM dbo.transactions
UNION ALL
SELECT 'agents', COUNT(*), COUNT(DISTINCT agent_id),
       SUM(CASE WHEN agent_id IS NULL THEN 1 ELSE 0 END)
FROM dbo.agents
UNION ALL
SELECT 'loans', COUNT(*), COUNT(DISTINCT loan_id),
       SUM(CASE WHEN loan_id IS NULL THEN 1 ELSE 0 END)
FROM dbo.loans
UNION ALL
SELECT 'repayments', COUNT(*), COUNT(DISTINCT repayment_id),
       SUM(CASE WHEN repayment_id IS NULL THEN 1 ELSE 0 END)
FROM dbo.repayments;
GO

-- A difference between rows and distinct IDs needs investigation, including
-- cases where the same ID has different values in the raw source.
SELECT 'Missing customer signup date' AS check_name, COUNT(*) AS affected_rows
FROM dbo.customers WHERE signup_date IS NULL
UNION ALL
SELECT 'Missing transaction date', COUNT(*)
FROM dbo.transactions WHERE txn_date IS NULL
UNION ALL
SELECT 'Missing transaction amount or fee', COUNT(*)
FROM dbo.transactions WHERE amount_ghs IS NULL OR fee_ghs IS NULL
UNION ALL
SELECT 'Missing loan disbursement date or principal', COUNT(*)
FROM dbo.loans WHERE disbursement_date IS NULL OR principal_ghs IS NULL
UNION ALL
SELECT 'Missing repayment due date or amount due', COUNT(*)
FROM dbo.repayments WHERE due_date IS NULL OR amount_due_ghs IS NULL;
GO

-- Identify nonblank foreign keys that do not have a matching parent.
SELECT 'Transactions with missing or unmatched customers' AS check_name, COUNT(*) AS affected_rows
FROM dbo.transactions AS t
LEFT JOIN dbo.customers AS c ON t.customer_id = c.customer_id
WHERE t.customer_id IS NULL OR c.customer_id IS NULL
UNION ALL
SELECT 'Transactions with unmatched nonblank agents', COUNT(*)
FROM dbo.transactions AS t
LEFT JOIN dbo.agents AS a ON t.agent_id = a.agent_id
WHERE t.agent_id IS NOT NULL AND a.agent_id IS NULL
UNION ALL
SELECT 'Loans with missing or unmatched customers', COUNT(*)
FROM dbo.loans AS l
LEFT JOIN dbo.customers AS c ON l.customer_id = c.customer_id
WHERE l.customer_id IS NULL OR c.customer_id IS NULL
UNION ALL
SELECT 'Repayments with missing or unmatched loans', COUNT(*)
FROM dbo.repayments AS r
LEFT JOIN dbo.loans AS l ON r.loan_id = l.loan_id
WHERE r.loan_id IS NULL OR l.loan_id IS NULL;
GO

-- Data-quality flag: source files include loans disbursed before customer
-- signup. Investigate and disclose this; do not invent replacement dates.
SELECT COUNT(*) AS loans_before_customer_signup
FROM dbo.loans AS l
INNER JOIN dbo.customers AS c ON l.customer_id = c.customer_id
WHERE l.disbursement_date < c.signup_date;
GO

/* =============================================================================
   SECTION D: EIGHT BUSINESS QUESTIONS
   Run and export each numbered question separately, including column headers.
   Each question below returns one result set.
   ============================================================================= */

-- BEGIN Q1: Monthly successful transaction performance in 2025.
SELECT
    CONVERT(char(7), t.txn_date, 120) AS transaction_month,
    COUNT(*) AS transaction_count,
    SUM(t.amount_ghs) AS total_value_ghs,
    COUNT(DISTINCT t.customer_id) AS active_customers
FROM dbo.transactions AS t
WHERE t.status = 'Success'
  AND t.txn_date >= '20250101'
  AND t.txn_date <  '20260101'
GROUP BY CONVERT(char(7), t.txn_date, 120)
ORDER BY transaction_month;
-- END Q1: export this result as q1.csv.
GO

-- BEGIN Q2: Channel success and failure rates; all statuses in denominator.
SELECT
    t.channel,
    COUNT(*) AS total_transactions,
    SUM(CASE WHEN t.status = 'Success'  THEN 1 ELSE 0 END) AS successful_transactions,
    SUM(CASE WHEN t.status = 'Failed'   THEN 1 ELSE 0 END) AS failed_transactions,
    SUM(CASE WHEN t.status = 'Reversed' THEN 1 ELSE 0 END) AS reversed_transactions,
    CAST(100.0 * SUM(CASE WHEN t.status = 'Success' THEN 1 ELSE 0 END)
         / NULLIF(COUNT(*), 0) AS decimal(5,2)) AS success_rate_pct,
    CAST(100.0 * SUM(CASE WHEN t.status = 'Failed' THEN 1 ELSE 0 END)
         / NULLIF(COUNT(*), 0) AS decimal(5,2)) AS failure_rate_pct,
    CAST(100.0 * SUM(CASE WHEN t.status = 'Reversed' THEN 1 ELSE 0 END)
         / NULLIF(COUNT(*), 0) AS decimal(5,2)) AS reversal_rate_pct
FROM dbo.transactions AS t
WHERE t.txn_date >= '20250101'
  AND t.txn_date <  '20260101'
GROUP BY t.channel
ORDER BY t.channel;
-- END Q2: export this result as q2.csv.
GO

-- BEGIN Q3: Top 10 agents by successful transaction value, followed by agents
-- with no 2025 transactions of any status. A message row is returned if none.
;WITH AgentStats AS (
    SELECT a.agent_id, a.agent_region, a.agent_type,
           COUNT(*) AS successful_transactions,
           SUM(t.amount_ghs) AS total_value_ghs
    FROM dbo.agents AS a
    INNER JOIN dbo.transactions AS t ON a.agent_id = t.agent_id
    WHERE t.status = 'Success'
      AND t.txn_date >= '20250101' AND t.txn_date < '20260101'
    GROUP BY a.agent_id, a.agent_region, a.agent_type
), RankedAgents AS (
    SELECT agent_id, agent_region, agent_type, successful_transactions,
           total_value_ghs,
           ROW_NUMBER() OVER (ORDER BY total_value_ghs DESC, agent_id) AS agent_rank
    FROM AgentStats
)
SELECT '1. Top 10 agents' AS result_section, agent_id, agent_region, agent_type,
       successful_transactions, total_value_ghs
FROM RankedAgents
WHERE agent_rank <= 10
UNION ALL
SELECT '2. Agents without transactions', a.agent_id, a.agent_region, a.agent_type,
       CAST(NULL AS int), CAST(NULL AS decimal(19,2))
FROM dbo.agents AS a
WHERE NOT EXISTS (
    SELECT 1 FROM dbo.transactions AS t
    WHERE t.agent_id = a.agent_id
      AND t.txn_date >= '20250101' AND t.txn_date < '20260101'
)
UNION ALL
SELECT '2. No agents without transactions', CAST(NULL AS nvarchar(100)),
       CAST(NULL AS nvarchar(100)), CAST(NULL AS nvarchar(100)),
       CAST(NULL AS int), CAST(NULL AS decimal(19,2))
WHERE NOT EXISTS (
    SELECT 1 FROM dbo.agents AS a
    WHERE NOT EXISTS (
        SELECT 1 FROM dbo.transactions AS t
        WHERE t.agent_id = a.agent_id
          AND t.txn_date >= '20250101' AND t.txn_date < '20260101'
    )
)
ORDER BY result_section, total_value_ghs DESC, agent_id;
-- END Q3: export this result as q3.csv.
GO

-- BEGIN Q4: Fee revenue from successful transactions by type and customer region.
SELECT
    t.txn_type AS transaction_type,
    COALESCE(c.region, 'UNMATCHED CUSTOMER') AS customer_region,
    COUNT(*) AS successful_transactions,
    SUM(t.fee_ghs) AS fee_revenue_ghs
FROM dbo.transactions AS t
LEFT JOIN dbo.customers AS c ON t.customer_id = c.customer_id
WHERE t.status = 'Success'
  AND t.txn_date >= '20250101' AND t.txn_date < '20260101'
GROUP BY t.txn_type, COALESCE(c.region, 'UNMATCHED CUSTOMER')
ORDER BY fee_revenue_ghs DESC, transaction_type, customer_region;
-- END Q4: export this result as q4.csv.
GO

-- BEGIN Q5: New customers per signup month and acquisition channel in 2025.
SELECT
    CONVERT(char(7), c.signup_date, 120) AS signup_month,
    c.acquisition_channel,
    COUNT(DISTINCT c.customer_id) AS new_customers
FROM dbo.customers AS c
WHERE c.signup_date >= '20250101' AND c.signup_date < '20260101'
GROUP BY CONVERT(char(7), c.signup_date, 120), c.acquisition_channel
ORDER BY signup_month, c.acquisition_channel;
-- END Q5: export this result as q5.csv.
GO

-- BEGIN Q6: Loans by product and by customer region in 2025.
SELECT 'Product' AS analysis_by,
       COALESCE(l.product, 'UNKNOWN PRODUCT') AS group_name,
       COUNT(*) AS number_of_loans,
       SUM(l.principal_ghs) AS total_lent_ghs,
       CAST(AVG(CAST(l.principal_ghs AS decimal(19,2))) AS decimal(19,2)) AS average_loan_size_ghs
FROM dbo.loans AS l
WHERE l.disbursement_date >= '20250101' AND l.disbursement_date < '20260101'
GROUP BY COALESCE(l.product, 'UNKNOWN PRODUCT')
UNION ALL
SELECT 'Region', COALESCE(c.region, 'UNMATCHED CUSTOMER'), COUNT(*),
       SUM(l.principal_ghs),
       CAST(AVG(CAST(l.principal_ghs AS decimal(19,2))) AS decimal(19,2))
FROM dbo.loans AS l
LEFT JOIN dbo.customers AS c ON l.customer_id = c.customer_id
WHERE l.disbursement_date >= '20250101' AND l.disbursement_date < '20260101'
GROUP BY COALESCE(c.region, 'UNMATCHED CUSTOMER')
ORDER BY analysis_by, total_lent_ghs DESC, group_name;
-- END Q6: export this result as q6.csv.
GO

-- BEGIN Q7: On-time repayment percentage and average days late by product.
-- Unpaid/partially paid due items are included in the denominator. Payments
-- after the year-end snapshot are not counted as fully paid as of year-end.
SELECT
    l.product,
    COUNT(*) AS due_repayments,
    SUM(CASE WHEN r.paid_date IS NOT NULL
                  AND r.paid_date <= r.due_date
                  AND r.amount_paid_ghs >= r.amount_due_ghs
             THEN 1 ELSE 0 END) AS on_time_repayments,
    SUM(CASE WHEN r.paid_date IS NOT NULL
                  AND r.paid_date > r.due_date
                  AND r.paid_date < '20260101'
                  AND r.amount_paid_ghs >= r.amount_due_ghs
             THEN 1 ELSE 0 END) AS late_paid_repayments,
    CAST(100.0 * SUM(CASE WHEN r.paid_date IS NOT NULL
                               AND r.paid_date <= r.due_date
                               AND r.amount_paid_ghs >= r.amount_due_ghs
                          THEN 1 ELSE 0 END)
         / NULLIF(COUNT(*), 0) AS decimal(5,2)) AS on_time_pct,
    CAST(AVG(CASE WHEN r.paid_date IS NOT NULL
                        AND r.paid_date > r.due_date
                        AND r.paid_date < '20260101'
                        AND r.amount_paid_ghs >= r.amount_due_ghs
                  THEN CAST(DATEDIFF(day, r.due_date, r.paid_date) AS decimal(12,2))
             END) AS decimal(12,2)) AS avg_days_late
FROM dbo.repayments AS r
INNER JOIN dbo.loans AS l ON r.loan_id = l.loan_id
WHERE r.due_date IS NOT NULL AND r.due_date < '20260101'
  AND l.disbursement_date >= '20250101' AND l.disbursement_date < '20260101'
GROUP BY l.product
ORDER BY l.product;
-- END Q7: export this result as q7.csv.
GO

-- BEGIN Q8: Default rates by product, customer region and KYC tier.
-- Each group must contain at least 30 loans. Do not add totals across sections:
-- each loan appears once in each relevant analysis section.
;WITH LoanBase AS (
    SELECT l.product, l.status, c.region, c.kyc_level
    FROM dbo.loans AS l
    LEFT JOIN dbo.customers AS c ON l.customer_id = c.customer_id
    WHERE l.disbursement_date >= '20250101' AND l.disbursement_date < '20260101'
), LoanGroups AS (
    SELECT 'Product' AS analysis_by, COALESCE(product, 'UNKNOWN PRODUCT') AS group_name,
           COUNT(*) AS total_loans,
           SUM(CASE WHEN status = 'Defaulted' THEN 1 ELSE 0 END) AS defaulted_loans
    FROM LoanBase
    GROUP BY COALESCE(product, 'UNKNOWN PRODUCT')
    HAVING COUNT(*) >= 30
    UNION ALL
    SELECT 'Region', COALESCE(region, 'UNMATCHED CUSTOMER'), COUNT(*),
           SUM(CASE WHEN status = 'Defaulted' THEN 1 ELSE 0 END)
    FROM LoanBase
    GROUP BY COALESCE(region, 'UNMATCHED CUSTOMER')
    HAVING COUNT(*) >= 30
    UNION ALL
    SELECT 'KYC Tier', COALESCE(kyc_level, 'UNKNOWN KYC'), COUNT(*),
           SUM(CASE WHEN status = 'Defaulted' THEN 1 ELSE 0 END)
    FROM LoanBase
    GROUP BY COALESCE(kyc_level, 'UNKNOWN KYC')
    HAVING COUNT(*) >= 30
)
SELECT analysis_by, group_name, total_loans, defaulted_loans,
       CAST(100.0 * defaulted_loans / NULLIF(total_loans, 0) AS decimal(5,2)) AS default_rate_pct
FROM LoanGroups
ORDER BY analysis_by, default_rate_pct DESC, group_name;
-- END Q8: export this result as q8.csv.
GO

/* Reconcile Q1, Q2 and Q8 with the Power BI report before submission. */
