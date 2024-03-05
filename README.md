# Excel Personal Budget Project
This project showcases how the power of Excel can be leveraged for a personal budget. Using tables, visualizations, formulae (like SUMIF and VLOOKUP), sliders, dropdowns, VBA macros, and pivot tables with slicers, a user can get high level or detailed overviews of their finances.

Below is a description of what each sheet in the workbook contains.

## Summary
This sheet (Figure 1) provides a summary overview of the user's finances.

> ![Screen Shot 2024-01-25 at 2 45 20 PM (2)](https://github.com/nwferreri/excel-budget-project/assets/112211174/146468b2-6327-4a84-82ae-e29f7f71d926)
>
> **Figure 1.** Summary sheet.

* The left column contains a welcome message that updates dynamically based on the present date.
* The numbers in the left column are dynamically calculated by referencing cells and tables in this and other sheets.
* The middle column visualizes a comparison between monthly income and expenses.
* The right column visualizes the relative amounts spent on different categories using data from another sheet.

## Savings Calculator
This sheet (Figure 2) allows the user to plan a savings goal.

> ![Screen Shot 2024-01-25 at 3 04 31 PM (2)](https://github.com/nwferreri/excel-budget-project/assets/112211174/269b0d16-4b06-4ca6-ae9f-73eba70ce7c5)
>
> **Figure 2.** Savings Goal Calculator.

* The savings goal and target date can be adjusted using the corresponding sliders.
* The user must update their current savings and the sheet will calculate how much they need to save each month to meet the goal.
* The user can choose whether or not to deduct the savings from their budget by using the dropdown on the right.

## Monthly Income & Monthly Expenses
These sheets contain tables of data that are used to calculate totals on the Summary sheet.

## Spending Summary
This sheet (Figure 3) contains a Pivot Table that summarizes the user's transaction history.

> ![Screen Shot 2024-01-25 at 3 47 32 PM (2)](https://github.com/nwferreri/excel-budget-project/assets/112211174/192b6d55-f899-480c-970f-fbe6d695c6cc)
>
> **Figure 3.** Spending summary pivot table.

* The total amount spent can be viewed by category or description across years and months.
* A slicer is provided, allowing the user to filter by the larger category groups.

## Transaction History
This sheet (Figure 4) contains the full list of the user's transactions.

> ![Screen Shot 2024-01-25 at 4 27 25 PM (2)](https://github.com/nwferreri/excel-budget-project/assets/112211174/902c4398-a358-4d02-9376-4b86a2e1ef57)
>
> **Figure 4.** Transaction history

* The final column is generated using a VLOOKUP to a table in another sheet.

## Other Sheets
The final 4 sheets contain various intermediate steps to generate the other aspects of the workbook.
* Raw Data: the user can paste in the raw transaction data from their financial institution and use a VBA macro to format it to match the Transaction History table.
* Chart Data: contains the data for the Expenses by Category chart in the Summary sheet. It uses SUMIFS functions to calculate the amount spent in each category.
* Behind the Scences: contains the validation list and scaling cells tied to the dropdown and sliders in the Savings Calculator sheet.
* Lookup Table: contains the lookup table for the VLOOKUP formula in the Transaction History sheet.
