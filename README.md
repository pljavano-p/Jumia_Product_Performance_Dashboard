# Building an Interactive Excel Dashboard for E-commerce Product Analysis: A Case Study of Jumia Products

## Introduction

Cleaning data and building dashboards in Excel is not just a classroom exercise; it is one of the most practical skills an analyst can have, since most business data still lives and gets reported in spreadsheets (Niklas, 2026). In this article, I take you through a real business problem: Jumia sellers need to understand how price, discounts, and customer reviews relate to product performance. I walk through cleaning a messy scraped dataset, writing formulas to enrich it, building PivotTables, and putting together an interactive dashboard that answers five specific business questions. By the end, you will know how to turn raw, inconsistent data into a dashboard that management can actually use.

## 1. The Dataset and the Problem
The dataset used for this project contains 115 Jumia product listings, scraped directly from the platform. It includes six columns: Product, Current price, old price, Discount, Review, and Rating, covering a mix of household, electronics, and personal care items. As shown in Figure 1, the data arrives in a raw, inconsistent state, exactly how real business data tends to look before anyone has touched it.

Before doing any cleaning, I went through the dataset and documented every quality issue I could find, rather than fixing things on the fly. Here is what stood out:

- Misspelled header: the Rating column header was actually written as "Ratingd" instead of "Rating."

- Negative review counts: values in the Review column appeared as negative numbers, for example, -2 or -14, even though a review count can never be negative.

- Price stored as a range: one row listed its price as "1,620 - 1,980" instead of a single number, breaking the pattern every other row followed.

- Rating stored as text: ratings were written as full phrases, such as "4.5 out of 5," rather than plain decimal numbers, making them unusable for calculations.

- Missing values: roughly half the rows had no Review or Rating values, leaving visible blanks throughout the dataset.

![raw_dataset.png](/screenshots/raw_dataset.png)

## 2. Cleaning the Data

Cleaning began with the Product column, applying `=PROPER(TRIM(A2))` to fix inconsistent spacing and capitalization, followed by Find and Replace to correct acronyms like USB and DIY that got wrongly lowercased. Prices were stripped of currency symbols and commas, review counts had their negative signs removed, and the one price range was resolved to a single value. Ratings were converted from text like "4.5 out of 5" into plain decimals using `=IFERROR(VALUE(LEFT(F2,FIND(" ",F2)-1)),"")`. Figure 2 shows the cleaned dataset.

![cleaned_dataset1.png](/screenshots/cleaned_dataset1.png)

## 3. Enrichment: Turning Raw Numbers into Categories
A rating of 4.5 or a discount of 42 percent doesn't mean much to a PivotTable on its own, since numbers like that vary slightly and are hard to group. To make the data easier to summarize, I added three enrichment columns: Rating Category, Discount Category, and Discount Amount, as shown in Figure 3. These take the raw numbers and sort them into simple buckets like "Excellent" or "High Discount," which is what actually makes counting, comparing, and charting possible later in the PivotTables and dashboard.

`=IF([@Rating]="","Missing",IF([@Rating]<3,"Poor",IF([@Rating]<=4.5,"Average","Excellent")))`

`=IF([@Discount]="","Missing",IF([@Discount]<20%,"Low Discount",IF([@Discount]<=40%,"Medium Discount","High Discount")))`

`=[@[old price (Ksh)]]-[@[Current price (Ksh)]]`

![cleaned_dataset2.png](/screenshots/cleaned_dataset2.png)
 
## 4. Building the Analysis Layer

With the data cleaned and categorized, the next step was turning it into actual answers. I built a separate Analysis sheet to hold three things: a set of key performance indicators, correlation checks between the main variables, and ranked Top 10 tables for rating, reviews, and discount. Figure 4 shows the KPI section, along with the correlations and the Top 10 by Rating table, providing a quick snapshot of the whole dataset: 115 products, an average price of about KSh 1,173, an average discount of 37 percent, an average rating of 3.9, and 723 total reviews.

![analysis1.png](/screenshots/analysis1.png)
 
The correlations, also visible in Figure 4, are what actually answer three of the five business questions:

`=CORREL(tblProducts[Discount],tblProducts[Review]) → -0.14`

`=CORREL(tblProducts[Rating],tblProducts[Review]) → 0.06`

`=CORREL(tblProducts[Current price (Ksh)],tblProducts[Rating]) → 0.11`


All three values sit close to zero, meaning none of these relationships are strong in this dataset, a finding that turns out to matter a lot once we get to the insights.

Building the Top 10 tables uncovered a bug worth mentioning on its own. Several products shared the same rating, so a simple ranking formula kept returning the same product over and over instead of listing ten different ones. I fixed this by building a "Rank Key" for each ranking, a small formula that nudges each value by a tiny, unique amount based on review count and row number, just enough to break ties without changing the actual order. Figure 5 shows the Top 10 by Reviews and Top 10 by Discount tables, both correctly listing ten distinct products once the fix was applied.

![analysis2.png](/screenshots/analysis2.png)

## 5. PivotTables: Summarizing for the Dashboard

To turn the analysis into something chartable, I built six PivotTables, shown together in Figure 6. Discount Mix and Rating Mix count how many products fall into each category. Engagement by Discount shows average reviews per discount tier. Top Products by Rating, Reviews, and Discount rank the ten best performers in each area. The Discount vs Rating Cross-tab combines both categories, revealing that 22 products have a high discount and only an average or poor rating.

![pivot_tables.png](/screenshots/pivot_tables.png)
 
## 6. Building the Dashboard

Everything built so far comes together in one sheet, shown in Figures 7 and 8. It opens with six KPI cards, followed by three scatter charts, three bar charts ranking top products, and two doughnut charts showing category mixes. Slicers for Discount and Rating Category sit near the top, connected to every relevant PivotChart at once, so clicking one button updates several charts together instead of digging through each pivot separately. That connection is what actually makes this a dashboard rather than just a page of static charts.

![dashboard1.png](/screenshots/dashboard1.png)

![dashboard2.png](/screenshots/dashboard2.png)

## 7. What Broke and How I Fixed It

- Text-formatted cells break formulas: Typing `=PROPER(TRIM(A2))` into a cell formatted as Text caused Excel to store the formula as plain text instead of calculating it. I fixed this by changing the cell format to General and running Text to Columns to force Excel to recalculate.
  
- Duplicate products in ranked tables: Several products shared the same Rating and Review count, causing my INDEX and MATCH formulas to return the same product repeatedly instead of listing 10 different ones. I fixed this by building a composite tie breaker key that combined Rating, Review count, and row number, giving every product a unique value to rank by.
  
- Scatter chart axis assignment: Excel decides which column becomes the X-axis based on which one is further left on the sheet, not on which one we actually select first when building the chart. This caused my Rating versus Reviews chart to plot backwards, with Reviews on the X axis instead of Rating.
  
- SPILL error with FILTER: Dynamic array formulas like FILTER cannot spill their results inside an Excel Table. Once I moved the formula into a plain range outside the table, it worked correctly and populated as expected.

## 8. Key Findings

Figure 9 shows the final Key Insights panel from the dashboard, summarizing the answer to each business question:

- Larger discounts do not bring more reviews (correlation -0.14)

- Higher ratings do not drive more engagement (correlation 0.06)

- Price and rating barely relate (correlation 0.11)

- A small group of listings drives most engagement

- 22 products need pricing and marketing attention

![key_insights.png](/screenshots/key_insights.png)
 
## Conclusion

Working through this project took a messy, scraped dataset and turned it into a finished, interactive dashboard. Here is a summary of what was covered:

- Auditing and cleaning raw data, fixing text formatting, broken acronyms, negative values, and text stored where numbers belonged

- Writing formulas to enrich the data with categories and rankings

- Building PivotTables to summarize discount, rating, and engagement patterns

- Assembling an interactive dashboard with KPI cards, charts, and slicers

- Turning correlations and rankings into five clear, verdict-led business insights

## Project Files

[Excel Workbook](https://github.com/ochiengherman36/Jumia_Product_Performance_Dashboard/blob/main/Excel_jumia_dataset.xlsx)

## Reference

Niklas. (2026). Why excel is still the backbone of business reporting in the AI era. Learnesy. https://learnesy.com/excel-is-still-the-backbone-of-business-reporting/
