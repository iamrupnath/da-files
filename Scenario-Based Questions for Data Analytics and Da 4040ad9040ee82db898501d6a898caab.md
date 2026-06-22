# Scenario-Based Questions for Data Analytics and Data Engineering Interviews

# This guide is designed for candidates preparing for Data Analyst, Data Engineer, Business Analyst, Analytics Engineer, and similar data-domain interviews.

The focus is not on basic definitions. The focus is on real-world scenarios where the interviewer wants to check whether the candidate understands data quality, business logic, SQL thinking, pipeline design, validation, debugging, and production-level problem solving.

A strong candidate should not only know the tool. A strong candidate should be able to explain the approach, assumptions, edge cases, and final solution clearly.

---

# Question 1: Revenue Dropped by 40% Overnight. How Would You Investigate?

## Scenario

The business team reports that revenue has dropped by 40% compared to the previous day. The dashboard is showing a sudden decline, and leadership wants an explanation.

## Detailed Answer

The first thing I would do is avoid assuming that the business actually lost 40% revenue. In data roles, sudden metric changes can happen because of real business reasons, but they can also happen because of data issues, pipeline failures, missing files, tracking changes, or incorrect logic in the reporting layer.

I would first try to separate the problem into two possibilities: either it is a real business drop, or it is a data reporting issue. To validate this, I would start from the source system and move step by step through the data pipeline.

First, I would check whether the source data arrived completely. For example, if the revenue dashboard depends on orders data, I would compare today's order count with the previous few days. If the number of orders received today is much lower than usual, then the issue may be with data ingestion rather than actual business performance.

Next, I would verify whether the ETL pipeline completed successfully. I would check pipeline logs, job status, failed tasks, delayed files, rejected records, and whether all partitions were processed. If the pipeline only processed partial data, the dashboard may show a false revenue drop.

After that, I would validate the transformation logic. I would check whether any recent code change affected revenue calculation. Sometimes a new filter, incorrect join, changed date condition, currency conversion issue, or duplicate removal rule can impact the final number.

Once the data pipeline looks correct, I would break the revenue down by dimensions such as region, product category, payment method, channel, and customer segment. This helps identify whether the drop is happening everywhere or only in one specific area. If one region or one payment method is responsible for the decline, the issue becomes easier to investigate.

Finally, I would compare the dashboard number with raw source-level aggregates. If the source system shows normal revenue but the dashboard shows a drop, then it is a data pipeline or reporting issue. If both source and dashboard show the drop, then it is likely a real business issue.

A strong answer here shows that I do not jump to conclusions. I validate the data first, then analyze the business reason.

---

# Question 2: Source Has 10 Million Records, Target Has 9.8 Million Records. How Would You Debug?

## Scenario

You are working on an ETL or migration project. The source table has 10 million records, but after loading into the target system, only 9.8 million records are available.

## Detailed Answer

I would not immediately treat this as a failed pipeline because sometimes the difference may be expected due to filters, duplicate removal, invalid records, or business rules. The first step is to understand whether the target is supposed to have the exact same count as the source.

If the expectation is one-to-one migration, then I would start validating the record count at every stage of the pipeline. For example, if the pipeline follows Bronze, Silver, and Gold layers, I would compare the count at source, Bronze, Silver, and Gold. This helps identify where the record loss happened.

If the count is correct in Bronze but reduced in Silver, then the issue is likely in cleaning, filtering, deduplication, or transformation logic. If the count is already reduced in Bronze, then the issue is likely in ingestion.

Next, I would compare counts by business date, region, category, or partition. Total count comparison is not enough because it only tells us that something is missing, not where it is missing. If a specific date has fewer records in target, I can focus my investigation there.

After that, I would perform key-level comparison. If the table has a primary key or business key, I would identify which keys exist in source but not in target. This can be done using anti-join logic. Once I get missing keys, I would inspect sample records to understand whether they were rejected, filtered, duplicated, or failed during transformation.

I would also check if any records were rejected due to data type mismatch, null values, invalid dates, schema mismatch, or constraint violations. In production systems, many records are lost silently because bad records are moved to reject tables or error logs.

Finally, I would not validate only counts. I would also compare important aggregates such as total revenue, total quantity, distinct customer count, and hash totals if needed. Count matching alone does not guarantee data correctness.

This answer shows that I understand real migration validation, not just simple row counting.

---

# Question 3: A Dashboard Shows Duplicate Customers. How Would You Find the Root Cause?

## Scenario

A business dashboard is showing the same customer multiple times. The business team expected one row per customer, but duplicates are appearing in the final report.

## Detailed Answer

The first thing I would check is the expected grain of the dataset. In data projects, duplication problems often happen because the grain is not clearly understood. If the final dashboard is expected to show one row per customer, then every transformation and join should preserve that customer-level grain.

I would start by checking whether duplicates already exist in the source data. I would group by customer ID and count the number of records. If the source itself contains multiple records per customer, then the dashboard may not be wrong. The issue may be that the business expected customer-level data, but the source is transaction-level or event-level data.

If the source is clean, then I would inspect the joins used in the transformation. Many duplicate issues happen because of one-to-many joins. For example, if a customer table is joined with an orders table, one customer can have many orders. If the query does not aggregate orders before joining, the customer record will multiply.

I would also check dimension tables. If a dimension table has multiple active records for the same customer due to Slowly Changing Dimension issues, then joining with that dimension can duplicate records.

Another important thing is to check join conditions. If the join is done only on customer name instead of customer ID, or if the join condition is incomplete, it can create incorrect matches and duplicate rows.

To fix the issue, I would first identify where duplication starts in the pipeline. Then I would either deduplicate the source, aggregate the many-side table before joining, fix the join condition, or apply correct SCD logic depending on the root cause.

A strong candidate should explain that duplicate rows are not always caused by duplicate data. They can also be caused by incorrect grain, joins, or modeling issues.

---

# Question 4: How Would You Detect Duplicate Transactions?

## Scenario

You work for a payment company. The business wants to identify duplicate transactions. A transaction may be duplicate if the same customer paid the same merchant the same amount within a short time window.

## Detailed Answer

Before writing any SQL, I would first clarify the business definition of duplicate transaction. In real projects, duplicate does not always mean two rows are exactly the same. Two transactions can have different transaction IDs but still be suspicious if the same customer paid the same amount to the same merchant within a few minutes.

I would ask the interviewer or business team what time window should be considered suspicious. For example, should we consider transactions duplicate if they happen within one minute, five minutes, or ten minutes? This definition is important because the query depends on it.

Once the definition is clear, I would partition the data by customer ID, merchant ID, and amount. Within each group, I would sort transactions by transaction time. Then I would use the LAG function to compare each transaction with the previous transaction for the same customer, merchant, and amount.

If the time difference between the current transaction and the previous transaction is less than or equal to the defined threshold, I would flag it as a possible duplicate.

The important part is that I would not delete or reject these records directly. I would first create a suspicious duplicate flag. Payment data is sensitive, and sometimes a customer may genuinely make two payments of the same amount. So the output should support investigation rather than blindly removing data.

This answer shows that I understand both SQL logic and business risk.

---

# Question 5: How Would You Calculate Customer Retention?

## Scenario

A product manager asks you to calculate monthly customer retention. You have an orders table with customer ID and order date.

## Detailed Answer

To calculate retention correctly, I would first define what retention means in this business context. Usually, retention means customers who performed an activity in one period and returned in a later period. In an e-commerce example, the activity could be placing an order. In an app, it could be opening the app or completing a transaction.

For monthly retention, I would first identify each customer's first purchase month. This first purchase month becomes the customer's cohort month. For example, if a customer placed their first order in January, then they belong to the January cohort.

Next, I would identify all later months in which the same customer placed orders. Then I would calculate the difference between the activity month and the cohort month. Month 0 means the first purchase month, Month 1 means the next month, Month 2 means two months later, and so on.

After that, I would count how many customers from each cohort returned in each later month. The retention rate would be calculated as the number of returning customers in a given month divided by the total number of customers in the original cohort.

A common mistake is counting orders instead of customers. Retention should usually be calculated at customer level, not order level. If one customer places ten orders in the next month, that is still one retained customer, not ten retained customers.

I would also be careful with incomplete months. If the current month is not completed, retention may look artificially low. So I would either exclude incomplete periods or clearly mention that the latest month is partial.

This answer shows that I understand metric design, cohort logic, and business interpretation.

---

# Question 6: A Pipeline Failed After Processing 80% of the Data. What Would You Do?

## Scenario

A daily ETL pipeline processes millions of records. One day, the job fails after processing 80% of the data. The business needs the final table urgently.

## Detailed Answer

In this situation, I would first avoid blindly rerunning the entire pipeline unless the pipeline is designed to safely handle reruns. If I rerun everything without understanding the current state, I may create duplicates, overwrite correct data, or increase processing cost unnecessarily.

The first thing I would check is whether the pipeline is idempotent. Idempotent means that running the same job multiple times should not create incorrect results. If the pipeline uses proper merge logic, partition overwrite, or delete-and-reload strategy for a specific date, rerunning may be safe.

Next, I would check where the failure happened. If the pipeline processes data by date, partition, batch ID, or file, I would identify which batches were completed and which were not. If checkpointing or watermarking is available, I would resume processing from the last successful checkpoint.

I would also check whether partial data has already been written to the target table. If partial records are present, I need to decide whether to clean them first or continue safely. For example, if the target supports merge based on primary key, I can continue processing without creating duplicates. But if the pipeline appends data blindly, I may need to delete the partially loaded partition and rerun that partition.

After solving the immediate issue, I would improve the pipeline design so that future failures are easier to recover from. This includes proper logging, checkpointing, batch tracking, retry logic, failure alerts, and reject record handling.

A strong answer should show production thinking. In real data engineering, failure is expected. The question is whether your pipeline can recover safely.

---

# Question 7: How Would You Validate Data After Migrating from One System to Another?

## Scenario

Your company migrated data from Snowflake to Databricks, or from an old data warehouse to a new platform. You need to prove that the migration is successful.

## Detailed Answer

I would not validate migration only by comparing total row counts. Row count comparison is important, but it is only the first level of validation. A table can have the same number of rows in source and target but still have incorrect values.

I would start with basic count validation at table level. Then I would compare counts by partition, date, region, category, or any important business dimension. This helps identify whether the mismatch is concentrated in a specific area.

Next, I would compare important aggregates. For example, in an orders table, I would compare total revenue, total order quantity, distinct customers, distinct orders, minimum and maximum dates, and null counts for important columns.

After aggregate validation, I would perform key-level validation. If the table has a primary key, I would check whether all source keys exist in target and whether all target keys exist in source. This helps identify missing or extra records.

For deeper validation, I would use row-level hash comparison. I would create a hash using important columns in the source and target, then compare hashes for matching primary keys. This helps identify records where the row exists in both systems but the values are different.

I would also validate data types, date formats, decimal precision, null handling, timezone handling, and special characters because these issues commonly happen during migration.

Finally, I would validate business reports generated from both systems. If the technical validation passes but business reports do not match, then there may be semantic differences in logic, filters, or modeling.

This answer shows practical migration knowledge and not just surface-level validation.

---

# Question 8: A Spark Job Is Running Very Slowly. How Would You Debug It?

## Scenario

A Spark job that usually finishes in 20 minutes is now taking more than two hours. You need to investigate and optimize it.

## Detailed Answer

I would start by opening the Spark UI because it gives the most practical view of what is happening inside the job. I would first check which stage is taking the most time. If one stage is much slower than others, I would inspect the tasks inside that stage.

If most tasks finish quickly but one or two tasks take a very long time, then the issue is likely data skew. Data skew means some partitions contain much more data than others. This usually happens when one key has a very high number of records, such as a popular customer, unknown category, or null key.

If all tasks are slow, then I would check shuffle size, input data size, number of partitions, executor memory, and whether the job is spilling data to disk. Shuffle is expensive because data moves across executors over the network. Large shuffle operations can slow down Spark jobs significantly.

I would also inspect the query logic. I would check whether filters are applied early, whether unnecessary columns are selected, whether joins are causing large shuffle, and whether a small table can be broadcasted.

If the issue is data skew, I may use salting, broadcast joins, skew join optimization with AQE, or repartitioning. If the issue is small files, I would compact files. If the issue is too many partitions or too few partitions, I would tune partition strategy.

I would also check whether the input data volume has increased recently. Sometimes the code is not the problem; the data volume has simply grown beyond the previous cluster configuration.

A strong answer should show that Spark optimization is not random configuration tuning. It starts with observation, then root cause analysis, then targeted optimization.

---

# Question 9: API Data Load Failed in the Middle. How Would You Design It Properly?

## Scenario

You are ingesting data from a paginated API. The API returns 1,000 records per page. The load fails after 300 pages because of a timeout.

## Detailed Answer

For API ingestion, I would design the pipeline assuming that failures can happen. APIs can fail because of timeout, rate limits, network issues, authentication expiry, or server-side errors.

First, I would implement pagination properly. The script should track which page, offset, or cursor has been processed. Instead of treating the API load as one large operation, I would process it in smaller batches.

Second, I would implement retry logic. If a request fails due to a temporary issue, the pipeline should retry after a delay. For rate limits, I would use exponential backoff or respect the retry-after header if the API provides one.

Third, I would maintain a checkpoint or state table. This table should store the last successfully processed page, timestamp, cursor, or batch ID. If the job fails, it should resume from the last successful checkpoint instead of starting from the beginning.

Fourth, I would ensure idempotency. If the same page is loaded again, it should not create duplicate records. This can be handled using primary keys, merge logic, or deduplication before writing to the final table.

I would also log failed requests, response codes, processing time, and record counts per page. This makes debugging easier when something goes wrong.

A strong API ingestion design should not fail silently and should not require manual guessing. It should be restartable, trackable, and safe to rerun.

---

# Question 10: How Would You Handle Slowly Changing Customer Data?

## Scenario

A customer changes their city from Bangalore to Mumbai. The business wants to preserve both old and new values for historical reporting.

## Detailed Answer

If the business wants to preserve history, I would use Slowly Changing Dimension Type 2. In this approach, we do not overwrite the old customer record. Instead, we create a new version of the record.

For example, the old record may show customer city as Bangalore with an effective start date and end date. When the customer moves to Mumbai, I would mark the old Bangalore record as inactive by updating its end date and current flag. Then I would insert a new record with city as Mumbai, a new start date, null end date, and current flag as true.

This allows historical reports to remain accurate. If the customer placed an order while living in Bangalore, the report should show Bangalore for that time period. If the customer places a new order after moving to Mumbai, the report should show Mumbai.

The most important part is identifying whether a change is meaningful enough to create a new version. Not every column change may require history tracking. For example, correcting spelling in a name may not need SCD Type 2, but changing city, customer segment, plan type, or account status may require it depending on business needs.

I would also handle late-arriving data carefully. If an old transaction arrives after the customer has changed city, the join should still map that transaction to the correct customer version based on effective dates.

This answer shows that I understand historical reporting and dimensional modeling, not just update statements.

---

# Question 11: How Would You Design a Daily Data Quality Framework?

## Scenario

Your company wants to monitor data quality before data reaches business dashboards.

## Detailed Answer

I would design the data quality framework around the most important risks in the data pipeline. The goal is not to create random checks, but to catch issues that can impact business reporting.

First, I would add completeness checks. This means validating whether expected data has arrived. For example, if the orders table receives data daily, I would check whether today's partition exists and whether the record count is within an acceptable range.

Second, I would add uniqueness checks. If order ID is expected to be unique, duplicate order IDs should be flagged. Duplicate primary keys can break joins, reports, and downstream calculations.

Third, I would add null checks for critical columns. Not every column needs a null check, but columns like customer ID, order ID, transaction amount, and event timestamp are usually important.

Fourth, I would add validity checks. For example, revenue should not be negative unless refunds are represented that way. Order date should not be in the future. Email should follow a valid format if used for communication.

Fifth, I would add referential integrity checks. If orders contain customer IDs, those customers should exist in the customer dimension table.

Finally, I would create alerting and reporting. A data quality framework is only useful if failures are visible. The system should clearly show which check failed, how many records failed, and whether the pipeline should stop or continue.

A strong answer should show that data quality is not an afterthought. It is part of pipeline design.

---

# Question 12: CEO Asks Why Sales Dropped. How Would You Structure the Analysis?

## Scenario

The CEO asks why sales dropped this month. You have access to sales, customers, products, regions, and marketing data.

## Detailed Answer

I would approach this as a business investigation, not just a SQL query.

First, I would confirm the metric definition. Sales can mean gross sales, net sales, revenue after discounts, booked revenue, or delivered revenue. If the definition is unclear, different teams may report different numbers.

Once the metric is clear, I would compare the current month with the previous month and the same month last year. This helps separate seasonal effects from actual decline.

Then I would break the decline across major dimensions. I would analyze sales by region, product category, customer segment, sales channel, and time period. The goal is to isolate the main driver of the decline.

For example, if total sales dropped by 20%, but one region dropped by 60%, then the issue may be regional. If one product category dropped significantly, then maybe inventory, pricing, or demand changed. If new customer sales dropped but existing customer sales stayed stable, then the issue may be acquisition.

After identifying the major contributor, I would go deeper. If the decline is in one channel, I would check traffic, conversion rate, average order value, cancellations, refunds, and marketing spend.

I would also check whether this is a data issue. If tracking changed, orders are delayed, or some source files are missing, the reported decline may not be real.

Finally, I would present the finding in a business-friendly way. Instead of saying, “Sales dropped because category X declined,” I would quantify the impact and explain the likely cause.

A good analysis should answer three things: what changed, where it changed, and why it likely changed.

---

# Question 13: Your Dashboard Numbers Do Not Match the Source System. How Would You Investigate?

## Scenario

The business team says the sales dashboard is showing different numbers compared to the source application. The source application shows 1.2 crore revenue, but the dashboard shows 1.05 crore.

## Detailed Answer

I would first avoid assuming that either the dashboard or the source system is wrong. In many cases, the difference happens because both systems are using different definitions, different time zones, different filters, or different stages of data.

The first thing I would do is confirm the metric definition. I would ask whether revenue means gross revenue, net revenue, revenue after discount, revenue after tax, or revenue after cancellations and refunds. Many dashboard mismatches happen because the source system shows gross numbers while the reporting system shows cleaned or adjusted numbers.

Next, I would check the date logic. If the source system uses transaction time and the dashboard uses order completion time or delivery date, numbers will not match. I would also check the timezone. For global or multi-region systems, UTC versus local timezone can create daily mismatches.

After that, I would compare the data at a lower grain. Instead of comparing only total revenue, I would compare revenue by date, product category, region, payment method, and order status. This helps identify where the mismatch is coming from.

I would also validate whether all source records reached the reporting layer. I would compare record counts between the source, raw layer, transformed layer, and dashboard layer. If the records are present in raw but missing in final reporting, then the issue is likely in transformation logic.

Another important step is checking filters. The dashboard may exclude cancelled orders, test orders, failed payments, internal users, or refunded transactions. If the source system includes them, the numbers will naturally differ.

Finally, I would document the finding clearly. If the mismatch is expected due to business logic, I would explain it with examples. If it is a data issue, I would identify the exact layer and logic causing the problem.

A strong answer shows that dashboard mismatch is not always a technical bug. Sometimes it is a metric definition issue.

---

# Question 14: How Would You Handle Late-Arriving Data in a Data Pipeline?

## Scenario

Your pipeline runs every day at 2 AM. However, some transactions from yesterday arrive late after the pipeline has already completed.

## Detailed Answer

Late-arriving data is very common in real data pipelines. It can happen because of delayed source systems, network issues, retries, vendor delays, or event processing delays. A strong data pipeline should be designed assuming that late data can arrive.

First, I would understand the business requirement. Some businesses need exact historical correction, while others are fine with approximate reporting after a cut-off time. For example, financial reporting usually requires correction, but some operational dashboards may accept minor delays.

If late-arriving records need to be included, I would avoid processing only the current date blindly. Instead, I would use a lookback window. For example, if late data usually arrives within three days, every daily run can reprocess the last three days. This ensures that delayed records are picked up.

The important part is making the pipeline idempotent. If I reprocess the last three days, I should not duplicate records in the target table. I would use merge logic based on primary key or business key, or overwrite specific date partitions safely.

For event-based pipelines, I would use watermarking to define how long the system should wait for late data. In streaming systems, watermarking helps manage late events while controlling state size.

I would also maintain ingestion timestamp and event timestamp separately. Event timestamp tells when the business event actually happened, while ingestion timestamp tells when the system received the record. This distinction is very important for debugging and reporting.

A strong answer should show that late-arriving data is not an exception. It is a normal design consideration in production pipelines.

---

# Question 15: How Would You Find Customers Who Have Not Purchased in the Last 90 Days?

## Scenario

The marketing team wants a list of customers who were active earlier but have not placed any order in the last 90 days.

## Detailed Answer

I would first clarify the business definition of inactive customer. A customer who never purchased and a customer who purchased earlier but stopped purchasing are different categories. For this problem, we are looking for customers who had at least one purchase in the past but no purchase in the last 90 days.

The simplest approach is to calculate the latest order date for each customer. I would group the orders table by customer ID and find the maximum order date. Then I would compare that latest order date with the current date. If the latest order date is older than 90 days, the customer can be considered inactive or churn-risk depending on the business definition.

I would also be careful with customers who placed cancelled or failed orders. If the orders table contains all order statuses, I would include only completed or valid orders in the calculation. Otherwise, a failed transaction may incorrectly make a customer appear active.

Another important point is customer lifecycle. If a customer purchased only once two years ago, they may not be useful for the same campaign as someone who purchased regularly and stopped recently. So, in a real analysis, I would also calculate total lifetime orders, lifetime value, and previous purchase frequency to prioritize customers.

This answer shows that I can write the SQL logic, but I also understand the business context behind customer inactivity.

---

# Question 16: How Would You Build a Monthly Revenue Report from Raw Orders Data?

## Scenario

You have a raw orders table containing order ID, customer ID, order date, amount, discount, tax, order status, and payment status. The business wants a monthly revenue report.

## Detailed Answer

I would first clarify what revenue means for the business. This is important because raw amount may not represent final revenue. Revenue could be calculated before discount, after discount, after tax, or only after payment is successful.

Once the definition is clear, I would filter only valid orders. For example, I may include only orders where payment status is successful and order status is completed or delivered. I would exclude failed, cancelled, test, or duplicate orders if the business rule says so.

Next, I would derive the month from the order date. I would ensure the date is based on the correct business timestamp. Some businesses report revenue based on order placed date, while others report based on delivery date or invoice date.

Then I would aggregate revenue at month level. I would also include supporting metrics such as order count, distinct customer count, average order value, discount amount, refund amount, and net revenue if available. This makes the report more useful than a single revenue number.

I would also validate the report by comparing monthly totals with source reports or finance reports. Revenue is a sensitive metric, so it should not be published without validation.

A strong answer shows that reporting is not only about group by month. It requires metric definition, filtering logic, date logic, and validation.

---

# Question 17: How Would You Identify Top Performing Products?

## Scenario

The business asks for the top performing products in the last quarter.

## Detailed Answer

I would first ask what “top performing” means. Many candidates assume it means highest revenue, but depending on the business, top performing could mean highest units sold, highest profit margin, highest repeat purchase rate, lowest return rate, or fastest growth.

If the business wants revenue-based ranking, I would calculate total revenue per product for the last quarter. But I would not stop there. I would also check quantity sold, number of unique customers, return rate, discount percentage, and margin if available.

A product may have high revenue only because it is expensive, but it may not sell frequently. Another product may have lower revenue but very high volume and better repeat behavior. Similarly, a product with high revenue and high return rate may not actually be good for business.

I would rank products using the agreed metric and then provide supporting context. For example, I could say, “Product A is highest by revenue, but Product B has the highest order volume and lower return rate.”

From a SQL perspective, I would filter the last quarter, aggregate at product level, and use ranking functions if we need top products by category or region.

A strong answer shows that I do not blindly rank by one metric. I clarify the business meaning first.

---

# Question 18: How Would You Detect a Sudden Spike in Failed Payments?

## Scenario

The payment failure rate suddenly increased from 3% to 18% in one day. The business wants to understand what happened.

## Detailed Answer

I would start by confirming whether the spike is real or caused by tracking or data issues. I would first compare total payment attempts, successful payments, and failed payments for the current day versus previous days.

Then I would break down the failure rate by payment method, bank, gateway, region, device type, app version, and time of day. Payment failures are usually not evenly distributed. The spike may be caused by one payment gateway, one bank, one app version, or one region.

I would also check error codes or failure reasons. If the majority of failures have the same error code, that gives a strong clue. For example, timeout errors may indicate gateway issues, while authentication errors may indicate customer-side or bank-side issues.

Next, I would check whether any release, configuration change, payment gateway routing change, or backend deployment happened before the spike. Many payment issues are caused by recent changes.

I would also check whether traffic increased suddenly. Sometimes failure count increases because total attempts increased, but failure rate gives the real picture. That is why both count and percentage are important.

A strong answer shows that I can investigate operational metrics by slicing the data and connecting it to business systems.

---

# Question 19: How Would You Handle Null Values in a Dataset?

## Scenario

You receive a customer dataset where important fields like email, phone number, city, age, and signup date contain null values.

## Detailed Answer

I would not use one common strategy for all null values. The right approach depends on the column, business use case, and downstream impact.

First, I would profile the null values. I would check how many nulls exist in each column and whether nulls are concentrated in specific sources, dates, regions, or customer segments. This helps identify whether the nulls are random or caused by a system issue.

For critical identifiers like customer ID or order ID, null values are usually not acceptable. Records with missing primary keys may need to be rejected or sent to a quarantine table because they cannot be joined or tracked reliably.

For optional fields like phone number or secondary email, nulls may be acceptable. In that case, I would keep them as null and ensure downstream reports handle them correctly.

For categorical fields like city or country, I may replace null with “Unknown” only if the reporting requirement needs a visible category. But I would be careful because replacing nulls can hide data quality issues.

For numerical fields like age or income, I would avoid blindly filling nulls with average values unless there is a strong analytical reason. Imputation can distort analysis if done incorrectly.

The most important thing is documenting the null handling rule. In production, null handling should be consistent and explainable.

A strong answer shows that null handling is not just a pandas fillna or SQL coalesce task. It is a data quality and business decision.

---

# Question 20: How Would You Handle Duplicate Records in a Data Pipeline?

## Scenario

A daily pipeline is loading duplicate records into the target table. The business is seeing inflated counts in reports.

## Detailed Answer

I would first identify whether the duplicates are coming from the source or being created inside the pipeline. To do this, I would check duplicate counts at each layer of the pipeline. If duplicates exist in the raw layer, the source is sending duplicates. If raw is clean but duplicates appear after transformation, then the issue is likely in joins, appends, or transformation logic.

Next, I would define what duplicate means. Exact row duplicates are easy to detect, but business duplicates may have different technical IDs. For example, the same transaction may appear twice with different ingestion timestamps. In that case, the duplicate definition should be based on business keys like customer ID, transaction amount, merchant ID, and transaction time.

If duplicates are caused by reruns, then the pipeline may not be idempotent. For example, if a failed job is rerun and the pipeline appends data again without deleting or merging previous records, duplicates will be created. To fix this, I would use merge/upsert logic, partition overwrite, or delete-and-reload for the affected batch.

If duplicates are valid in raw data but not in final reporting, I would keep raw data unchanged and deduplicate in the cleaned layer using ranking logic. For example, I could use row_number over the business key ordered by updated timestamp and keep the latest record.

A strong answer shows that duplicate handling requires root cause analysis. We should not blindly apply distinct because it can hide problems and remove valid records.

---

# Question 21: How Would You Design an Incremental Load?

## Scenario

You have a large source table with 500 million records. You need to load only new and changed records every day instead of reloading the entire table.

## Detailed Answer

For incremental loading, I would first identify the column that can help detect new or changed records. Usually, this is an updated_at timestamp, created_at timestamp, incrementing ID, or change data capture column.

If the source has a reliable updated_at column, I would maintain a watermark. The watermark stores the maximum updated_at value successfully processed in the previous run. In the next run, I would fetch only records where updated_at is greater than the last watermark.

However, I would be careful with late-arriving data. If some records arrive late with an older updated_at value, they may be missed. To handle this, I may use a small lookback window. For example, instead of loading only records after the exact watermark, I may reload the last one or two days and merge them into the target.

The target write should be idempotent. I would avoid blindly appending incremental data because changed records may already exist in the target. Instead, I would use merge logic based on the primary key. If the key exists, update the record. If it does not exist, insert it.

I would also maintain audit information such as batch ID, load timestamp, source count, target count, inserted count, updated count, and rejected count. This helps validate every incremental run.

A strong answer shows that incremental loading is not only about filtering on updated_at. It also requires watermarking, late data handling, merge logic, and auditing.

---

# Question 22: How Would You Explain the Difference Between Batch and Streaming in Real Projects?

## Scenario

An interviewer asks you when to use batch processing and when to use streaming.

## Detailed Answer

I would explain this using business requirements rather than only technical definitions.

Batch processing is useful when data does not need to be processed immediately. For example, daily sales reporting, monthly finance reports, customer segmentation, and historical analytics can usually run in batches. In batch pipelines, we collect data over a period and process it together.

Streaming is useful when the business needs data quickly, often within seconds or minutes. For example, fraud detection, payment monitoring, live order tracking, IoT alerts, and real-time dashboards may require streaming.

The key difference is not only speed. Streaming systems are usually more complex because they need to handle continuous data, late events, duplicates, state management, checkpointing, and fault tolerance.

I would also mention that many companies use a hybrid approach. For example, they may use streaming for operational alerts and batch processing for final financial reporting. Streaming gives speed, while batch gives stability and reconciliation.

A strong answer shows that I understand trade-offs. Real companies do not choose streaming just because it sounds advanced. They choose it when the business problem requires it.

---

# Question 23: How Would You Optimize a Slow SQL Query?

## Scenario

A SQL query used in a dashboard is taking 20 minutes to run. The business wants it to return results faster.

## Detailed Answer

I would first understand the query purpose and the tables involved. Then I would check the query execution plan to see where most of the cost is coming from. Without checking the execution plan, optimization becomes guesswork.

I would look for full table scans on large tables, expensive joins, missing filters, unnecessary columns, repeated subqueries, and incorrect join conditions. If the query reads many columns but uses only a few, I would select only the required columns.

Next, I would check whether filters are applied early. If a query joins huge tables first and filters later, it may process much more data than needed. Applying filters before joins can reduce the data volume significantly.

I would also check join keys. If joins are happening on non-indexed columns in an RDBMS, performance may be poor. If the join condition is incomplete, it may create row multiplication and increase processing time.

For aggregation queries, I would check whether pre-aggregated tables or materialized views can help. Dashboards should not always run heavy raw-level queries repeatedly.

I would also check data model design. Sometimes the problem is not the query but the fact that the dashboard is directly querying raw transactional tables instead of a properly designed reporting layer.

A strong answer shows that SQL optimization is not just adding indexes. It requires understanding execution plans, data volume, joins, filters, and reporting design.

---

# Question 24: How Would You Validate a Data Model Before Giving It to BI Team?

## Scenario

You have created a new reporting table for the BI team. Before handing it over, you need to validate it.

## Detailed Answer

I would validate the reporting table from both technical and business perspectives.

First, I would validate the grain of the table. If the table is supposed to have one row per customer per month, I would check whether that uniqueness holds. Many BI issues happen because the reporting table grain is unclear or broken.

Next, I would validate key metrics. I would compare totals from the new table with trusted source numbers. For example, revenue, order count, customer count, and quantity should match expected values within an acceptable range.

I would also validate dimensions. For example, region, product category, customer segment, and date fields should not have unexpected nulls or incorrect values.

Then I would test common BI use cases. I would run sample dashboard queries to ensure that filtering by date, region, category, or segment gives correct results.

I would also check performance. A reporting table should be easy for BI tools to query. If every dashboard query takes too long, the model may need aggregation, partitioning, clustering, or restructuring.

Finally, I would document the table definition, grain, metric logic, refresh frequency, known limitations, and owner. This helps BI users understand how to use the table correctly.

A strong answer shows that data modeling is not complete when the table is created. It is complete when the table is accurate, usable, performant, and documented.

---

# Question 25: How Would You Handle a Schema Change in Source Data?

## Scenario

A source system suddenly adds a new column, removes a column, or changes the data type of an existing column. Your pipeline fails.

## Detailed Answer

I would first check what kind of schema change happened. Adding a nullable column is usually less risky, while removing a column or changing a data type can break transformations and downstream reports.

If the pipeline fails because a required column is missing, I would check whether the source change was intentional or accidental. I would coordinate with the source team to understand the change and expected future schema.

For production pipelines, I would prefer having schema validation before transformation. The pipeline should detect unexpected schema changes and fail with a clear error message instead of producing incorrect data silently.

If the schema change is valid, I would update the pipeline logic, target table schema, data quality checks, and downstream documentation. If the removed column is used by reports, I would communicate impact to stakeholders before making changes.

For flexible ingestion layers, I may allow schema evolution in the raw or bronze layer, but I would be stricter in curated layers like silver and gold because these layers are consumed by downstream users.

A strong answer shows that schema changes should be controlled. Automatically accepting every schema change can be dangerous, especially in business-critical pipelines.

---

# Question 26: How Would You Monitor a Production Data Pipeline?

## Scenario

You have built a pipeline that runs every day and feeds business dashboards. How would you monitor it?

## Detailed Answer

I would monitor the pipeline at multiple levels: job execution, data quality, data freshness, and business metrics.

First, I would monitor whether the job ran successfully. This includes start time, end time, duration, failure status, retry status, and error messages. But job success alone is not enough. A pipeline can succeed technically and still produce wrong data.

Second, I would monitor data freshness. If the dashboard is expected to update by 8 AM, I would track whether the latest data is available on time. Freshness is very important for business trust.

Third, I would monitor volume checks. I would compare today's record count with historical patterns. If the pipeline usually processes 10 million records but today processed only 2 million, that should trigger an alert.

Fourth, I would monitor data quality checks such as nulls, duplicates, invalid values, referential integrity, and metric thresholds.

Fifth, I would monitor business-level metrics. For example, if daily revenue becomes zero or payment failure rate suddenly spikes, the system should alert even if the pipeline technically succeeded.

Finally, I would create clear alerts with useful details. A bad alert says “pipeline failed.” A good alert says which job failed, which table was affected, what the error was, and what business dashboard may be impacted.

A strong answer shows that monitoring is not only about job status. It is about trust in data.

---

# Question 27: How Would You Explain Data Granularity With an Example?

## Scenario

An interviewer asks why data granularity matters in analytics and data engineering.

## Detailed Answer

Data granularity means the level of detail represented by each row in a dataset. Understanding granularity is extremely important because many reporting errors happen when people join or aggregate data at the wrong level.

For example, suppose we have an orders table where each row represents one order. The grain is order-level. If we have an order_items table where each row represents one item inside an order, the grain is item-level. One order can have multiple items.

If I join orders with order_items and then calculate total order amount from the orders table, I may accidentally multiply revenue because one order row will repeat for every item. This is a common mistake in analytics.

Before writing any query, I should understand what one row represents. Is it one customer, one order, one transaction, one product per order, or one customer per month? Once the grain is clear, joins and aggregations become safer.

A strong answer shows that granularity is not a theoretical modeling term. It directly affects correctness of business metrics.

---

# Question 28: How Would You Find Why a Report Is Slow in Power BI, Tableau, or Any BI Tool?

## Scenario

A dashboard is taking too long to load. The business complains that every filter takes several seconds or minutes.

## Detailed Answer

I would investigate both the BI layer and the data layer. Sometimes the dashboard is slow because of the visualization design, and sometimes it is slow because the underlying query or model is inefficient.

First, I would check whether the dashboard is querying raw-level data directly. If the dashboard is running heavy joins and aggregations every time a user applies a filter, performance will be poor. In that case, I would create a pre-aggregated reporting table.

Next, I would check the number of visuals, filters, and calculated fields. Too many visuals on one page can trigger multiple queries. Complex calculated fields at BI level can also slow performance.

I would also check whether the data model has unnecessary many-to-many relationships or high-cardinality columns. High-cardinality filters, such as transaction ID or customer-level filters, can make dashboards slower.

At the database level, I would inspect the SQL queries generated by the BI tool. I would check whether filters are being pushed down properly, whether joins are efficient, and whether partitions or indexes are being used.

Finally, I would optimize by reducing data volume, creating proper aggregation tables, removing unnecessary columns, improving the data model, and shifting heavy calculations from BI layer to data warehouse where appropriate.

A strong answer shows that BI performance is a combined responsibility of dashboard design, data modeling, and backend optimization.

---

# Question 29: How Would You Handle Personally Identifiable Information in Data Pipelines?

## Scenario

Your pipeline contains customer email, phone number, address, and other sensitive information. How would you handle this data safely?

## Detailed Answer

I would first identify which columns contain sensitive or personally identifiable information. Common examples include email, phone number, address, government IDs, account numbers, and sometimes even location data.

Once identified, I would apply the principle of least privilege. Not every user or team should have access to raw sensitive data. Access should be role-based and limited to people who truly need it.

In raw layers, sensitive data may be stored securely, but in reporting layers, I would mask, hash, tokenize, or remove sensitive fields depending on the use case. For example, analysts may need customer-level behavior but not actual phone numbers or email addresses.

I would also ensure encryption at rest and in transit. Most modern cloud platforms provide this, but it still needs to be configured correctly.

Audit logging is also important. We should know who accessed sensitive data and when. This is especially important in regulated industries.

Another important point is data minimization. If a pipeline or report does not need phone number or address, we should not carry those columns forward unnecessarily.

A strong answer shows that I understand data privacy is not only a security team responsibility. Data engineers and analysts also play a critical role.

---

# Question 30: How Would You Explain the Difference Between OLTP and OLAP Using a Real Example?

## Scenario

An interviewer asks you to explain OLTP and OLAP in a practical way.

## Detailed Answer

I would explain OLTP as the system used to run daily business operations. For example, when a customer places an order on an e-commerce app, that transaction is stored in an OLTP database. The system needs to quickly insert, update, and retrieve individual records. It is optimized for small, fast transactions.

OLAP is used for analysis and reporting. For example, if leadership wants to know total sales by month, top products, customer retention, or region-wise revenue, that analysis is done on OLAP systems like data warehouses or lakehouses.

The same business can use both. The e-commerce application may use an OLTP database to process orders, payments, and customer details. That data is then moved into a data warehouse or lakehouse where analysts and data teams run reporting and analytics.

The key difference is workload. OLTP focuses on current operational transactions. OLAP focuses on historical analysis across large volumes of data.

A strong answer should connect OLTP and OLAP to real data flow instead of only saying one is transactional and the other is analytical.

---

# Final Note for Candidates

For every scenario, remember this answer structure:

First understand the business requirement. Then define the data grain and assumptions. After that, validate the data, identify possible causes, explain the solution, and mention edge cases.

Interviewers do not only want the final answer. They want to see whether you can think like someone who can work on real data systems.

---

# Final Note

In real interviews, the strongest candidates are not the ones who memorize the most definitions. The strongest candidates are the ones who can think clearly when the problem is messy.

For every scenario-based question, your answer should follow this pattern:

Understand the business requirement, define assumptions, validate the data, identify the root cause, explain the solution, and mention edge cases.

That is the difference between someone who knows tools and someone who knows how to work on real data problems.