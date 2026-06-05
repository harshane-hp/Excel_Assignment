                                                  Excel Data Manipulation & Analysis Practice

**Project Overview**

This repository contains an Excel spreadsheet used for practicing data entry, structuring, and fundamental formula operations. The dataset simulates a global retail sales log featuring product categorization, regional tracking, pricing, and inventory quantities across various months.

**I.**Dataset Structure****

The sheet consists of the following columns:

**Month**: The month the transaction occurred.

**Country Code**: Two-letter country code (e.g., US, UK, IN, CA) representing the market.

**Product Name**: The specific item sold (ranging from Electronics to Fashion and Kitchenware).

**Brand Name**: The manufacturer/brand of the product.

**Price ($)**: The individual cost of the item.

**Price Range**: Categorization of the pricing tier (e.g., High Price).

**Quantity**: The volume of units sold.

**Category**: The primary industry department (Electronics, Fashion, Kitchen, Outdoor, Accessories).


**II.**Performed Operations & Formulas****

**1.Sum:**

To determine the total volume of items handled across the recorded transactions, a summation function was implemented on the Quantity column data range.

Formula Used:
**=SUM(G2:G35)** and to add $ in front of the value **="$" & SUM(G2:G35)**. 

Answer is SUM=**$10100**

Purpose: Merges and totals all individual quantities from row 2 down to row 35 to provide a macro view of inventory turnover.

**2.Count:**

This function was used to determine the total number of transactions or entries listed in the dataset.

Formula:
**=COUNT(G2:G35)** .You'll get **34**.

Answer is COUNT=**34**

Purpose: Counts the number of cells containing numerical price data to verify the database.

**3.Average:**

To evaluate price points and find the central tendency of the inventory value

Formula Used:

**=AVERAGE(G2:G35)** and to add $ in front of the value **="$" & AVERAGE(G2:G35)**.You'll get  **$297.0588235**

To round off the decimal value for the provided average use **="$" & ROUND(AVERAGE(G2:G35),2)**

Now,the final Average is **$297.06**

**4.Minimum Price among all products**

To pinpoint the most budget-friendly product entry across all listed categories

Formula Used:
**=MIN(G2:G35)** and to add $ in front of the value **="$" & MIN(G2:G35)**. 

Answer is **$30**(T-Shirt)

Purpose:It instantly extracts the lowest baseline cost in your dataset without requiring you to manually sort the rows.

**5.Maximum Price among all products**

This function identifies the most expensive item across all product categories in the inventory log.

Formula Used:
**=MAX(G2:G35)** and to add $ in front of the value **="$" & MAX(G2:G35)**. 

Answer is **$1000**(Laptop)

Purpose:The primary purpose of the MAX function in Excel is to instantly find the highest numerical value within a specific range of cells. Instead of scrolling through hundreds of rows to spot the largest number, MAX does it mathematically in a split second.


**6.Logical Functions-IF CASE:**

To programmatically categorize each item into a premium or baseline pricing tier based on a $500 threshold

Formula Used:
**=IF(G2>=500, "Higher Price", "Standard Price")**

Purpose of the IF Function:

Automates Categorization: Instead of manually typing out pricing tiers for all 34 rows, the formula instantly standardizes the labels across the entire sheet.

Streamlines Data Filtering: It creates clean text markers ("Higher Price" vs "Standard Price") that allow you to easily filter your dashboard or create Pivot Tables to compare luxury items against budget ones.

Scalability: If a product's price updates in Column E, the classification label in Column F automatically shifts without requiring manual intervention.

**7**.**Conditional Functions – SUMIF and COUNTIF**

SUMIF:

To add up the total cost of all items that specifically belong to the Electronics department

Formula Used:
**=SUMIF(J2:J35,"Electronics",G2:G35)** and to add $ in front of the value **="$" & SUMIF(J2:J35,"Electronics",G2:G35)**

Answer is **$8050**

Purpose:It will skip other products and calculate the sum of electronics price 

COUNTIF:

To find out exactly how many individual product entries have a price point strictly below $100

Formula Used:**=COUNTIF(G2:G35, "<100")**

Answer is **11**

**8.Text Formatting – LEFT, RIGHT, MID**

**a.LEFT:**

The LEFT function pulls a specified number of characters starting from the very beginning (left side) of a text string.

Formula Used:=**LEFT(A2,2)**

(Extracting from the Beginning)
Purpose: To grab a specific number of characters starting from the very first character on the left side of a text cell.

Common Use Cases: * Extracting area codes from full phone numbers (e.g., pulling 415 from 415-555-0199).

Isolating the day from a standardized timestamp string.

Extracting the department code from a combined corporate account number.

**b.MID:**

Targets the middle segment of the text string, starting at character position 4 and extracting a length of 3 characters to isolate the Month identifier.

Formula Used:=**MID(A2,4,3)**

(Extracting from the Middle)

Purpose: To extract a specific chunk of text from anywhere inside the string, given a defined starting position and a character count.

Common Use Cases:
Extracting mid-string identifiers like a 3-letter month identifier (e.g., pulling AUG out of 23-AUG-2026).
Pulling out specific batch numbers or category designations hidden inside factory barcodes.

**c.RIGHT:**

Isolates the final 2 characters at the end of the string to pull the geographical Country Code

Formula Used:=**RIGHT(A2,2)**

(Extracting from the End)
Purpose: To grab a specific number of characters starting from the absolute last character on the right side of a text cell.

Common Use Cases:

Extracting regional or country codes located at the end of an order ID (e.g., pulling US or CA).

Isolating the last 4 digits of a credit card or bank account number for security verification displays.

Slicing off file extensions (like .jpg or .pdf) from a list of filenames.

                                                            **THANK YOU**


