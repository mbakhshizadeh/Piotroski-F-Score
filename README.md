# Piotroski-F-Score
Determination of the best-performing stocks using Piotroski F-Score 
Piotroski F-Score Rank model gives investors a ranking list of companies from top to bottom based on their performance in the current and last year. The important advantage of this method is that there is no need to have many data of the company (like startups). here, there needs onnly to have current and last year’s data (or quarters) to calculate the difference in the balance sheet components in the two sequential statements.
To calculate measurements of Piotroski-F-Score, first, we need to pull out these variables from balance sheets and then put them into formulas.

Variables:
Net income, Total Assets at the beginning of the year, Cash flow from operation, total long-term debt, current assets, current Liabilities, issued common stock, profit, and sales.

The output of the code is a table that shows the ranking list of companies based on their performance such that the first rank means the best-performing company.
