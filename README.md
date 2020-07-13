
## I. About

This project is provides a set of excel template files designed to make easier the process of creating and maintaining a budget, tracking expenses, and logically partitioning savings. The following document will describe how best to use these files. These templates are better suited to those with a predictable or fixed income, however may be used by those with a variable income with some tweaking to the methodologies described below.

## II. Basics

Each template file contains three sheets: a paycheck analysis, budget analysis, and savings analysis. The contents of these sheets should be edited with varying frequencies, which will be expanded on later. Each sheet breaks down its analysis by month, and as such the budget files should be edited on a monthly basis at minimum. Aside from the initial setup of a budget sheet, this upkeep should be very quick and easy. As a general rule of thumb, all cells that expect input are colored light blue, but this will be expanded on in further sections. Finally, the reason for the varying template files is so that a budget can be easily created regardless of the current month. The only difference between the templates is that all monthly inputs before the month specified in the file name have been replaced with dashes. (null values) 

## III. Creating A Budget Manually Or For The First Time

To begin start by making a copy of your desired template file, which will be used to build your budget from, then open the file in Excel.

### Paycheck Analysis

The first sheet is the Paycheck Analysis page, which should for the most part mirror your paycheck. This will probably only need to be edited once per year, at the time you create your budget from this template, since every month your gross pay should be the same. The exceptions to this rule would be if you receive a pay rate increase/decrease, or switch jobs. Any non-recurring or random additional income may also be tracked, but this is accounted for separately on the Budget Analysis sheet.

Starting on row 1 there’s a light blue cell with the current year, 2020. This value is used to populate the yearly values throughout each sheet and should be edited to reflect the current year. 

Next begin to edit the light blue cells under the “Inputs” header. Notice how in parenthesis it says that only one cell should be edited per row. The column of cells you should edit in this area is dependent on the frequency with which you receive your pay checks. Starting in row three, set one cell to a Y to denote your pay period, which will be used to help populate certain cells throughout the various sheets. Next go down that same column and populate the various inputs such as your Gross pay rate (before taxes), and various flat rates that are deducted per paycheck for benefits such as medical and vision coverage. Editing these values should update the values displayed under the “Calculations” header.

Next locate the “Variable Rates” subsection under Benefits and input the percentage of your paycheck that is taken out per pay period in cells F27 and F28. If you have additional account deductions you may add the account names in cells C29 and C30, and populate the relevant percentage columns accordingly. Similarly, under the “Taxes” subsection, you should input the percentage of your paycheck deducted through the various tax categories. If this is not listed on your paycheck as a percentage, but as a flat rate, hovering over cell F34 should tell you how this may be calculated.

With that done you should be finished editing this sheet, likely for the rest of your days using this budget, and won’t have to worry about any upkeep here unless you get a pay rate increase/decrease or switch jobs.

### Budget Analysis

Probably the most useful and important of the three sheets, the Budget Analysis is what is used to create your budget, predict and track your spending, and determine how much money you should be saving. This should be edited roughly once per month, probably at the end of the month, but of course may be adjusted more frequently if needed. In general the cells you’re going to be editing here are those that are highlighted in blue under the “Recurring Values” header, and those highlighted in blue under the “Actual Monthly” values header. 

The cells under “Recurring Values” header are used to calculate expected monthly values under the “Actual Monthly Values” header. These cells will likely be edited very infrequently; only in the case that your expected recurring inflow/outflow value has changed from what it was previously. The cells under the “Actual Monthly Values” on the other hand, should be edited once per month as part of your monthly upkeep with the **actual** inflow/outflow value for that month. Personally, I’ve found it helpful to get into the habit of changing these cell’s colors to transparent after editing, as it helps me easily spot the current month the next time I view my budget, and more importantly remind me to update the cell values if I haven’t already.  

*Note: Even if the calculated inflow/outflow value for a specific cell under “Actual Monthly Values” is the same as the expected, you should still overwrite the cell contents with that hard-coded number as part of your upkeep. This is so that if cell values under “Recurring” change at any point while you use this budget, the the cells values under “Actual Monthly Values” for previous months won’t also change as a result of the formulas. If you do not overwrite these cells, any changes to “Recurring” (expected) values will cause “Actual Monthly Values” for previous months to be overwritten.*

#### Cash Inflows

Starting with the “Cash Inflows” section, if you’ve filled out the Paycheck Analysis sheet, one of the cells H10-M10 (depending on your pay period) should reflect the net pay from that sheet. This should also populate the cells in row 10 under “Actual Monthly Values”. 

If you have any additional recurring income that isn’t a part of your paycheck, such as collected rent if you own and lease property, money collected on a recurring basis from tutoring, etc. then that may be specified under the “Additional Recurring Income” subsection, by replacing the dashes in C14 and C15 with the name of the income source. For these income sources the expected recurring income amount should be specified in **one** of the cells in that row under the “Recurring Values” header, depending on the frequency with which you expect to receive that income. This should auto calculate the actual monthly value for each month across the row. Finally, the “One Time Income” sub section is used to specify income that is not recurring, such as money made via overtime, money received as a gift, fulfilling a contract, or any other unexpected income for that month. For these values populate the cell for that month with the income amount received. Optionally, adding a comment to the cell that states where the income came from might be useful if you’d like to maintain a greater level of detail on your budget.

As specified in the comment above, the reason that a lot of the calculated cells under the “Actual Monthly Values” header are blue, is because they should be replaced at the end of the month with the actual values, even if they are the same.  It’s important to get into the habit of overwriting these values.

At the end of this section, the Available Income cells denote how much spending money you have a result of your total income for that month after taxes, but before any other cash outflows.

#### Cash Outflows

This section is arguably the most important of all, as it represents all the money that will get taken out of your available spending money for that month as a result of expenses or savings. The input sections work pretty much the same way they do in the Cash Inflows section.

The “Recurring Bills” sub-section is where you should specify expected recurring expenses such as service memberships and bills. I have some default line items that represent bills most people will have in common, but these can of course be replaced with items more relevant to you if needed. Again you should only edit *one* cell under the “Recurring Value” header per row in this section. Most of the time bills recur on a monthly basis and you’ll use that column, but if a certain bill happens weekly or in some other time frame you may use the according cell for that row instead. Certain recurring bills vary per month from the expected value you’ve inputted under the “Recurring Values” header, such as electricity. This doesn’t always necessitate a change to the expected value cell, but will necessitate (as always) that you update the cell under the “Actual Monthly Value” after you’ve paid that bill.

The next two sub sections denote savings accounts. I personally like to maintain two separate savings accounts; one as a “restricted” account which contains my emergency fund and funds for non-recurring expenses that I expect to hit randomly throughout the year, and the other for superfluous savings such as a vacation fund.

The line items under the “Restricted Savings” sub-section represent logical partitions or categories of savings that all fall under the common umbrella of money reserved for expected future expenses. Notice that these line items are numbered. That’s because the line items that you specify in these cells will be used to create matching line items on the Savings Analysis sheet, which will be discussed later. Any time you want to add an additional logical partition for restricted savings it should be done here, not on the Savings Analysis sheet. As always, pick **one** cell under the “Recurring Values” header per row where you should specify the money you expect to need. I usually use the ANNUAL column for this section, but it’s possible you may use a different column depending on what line item you’re using. Based on the cell you choose, the cell under the “Actual Monthly Values” header will auto calculate how much money should be saved for that category per month. The sub-total for this “Restricted Savings” sub-section will show the total amount that you should transfer to savings in a particular month for restricted expenses.

The “Other Savings” sub section works exactly the same way as the “Restricted Savings” but is meant to represent categories of superfluous, or non-critical savings, such as a vacation fund. In this section I typically prefer to ignore the cells under the “Recurring Values” header and simply populate the “Actual Monthly Values” cell based on the amount of money I have left over at the end of a particular month.

The last sub-section on this sheet is the “Spending Money” sub section. This section is useful for planning and tracking how much money you spend on things such as groceries, entertainment, and other variable expenses. By the time you get to this section you should have your various savings and recurring expenses sub sections filled out, which will have updated the “Excess Available Funds” cell for the month, listed below the “Cash Outflows” section. The amount listed in that cell represents the remaining funds that can be used on the various line items in this section. As you fill out expected values the “Excess Available Funds” cell value will update, so as long as that value is at or above 0 by the time you finish, you’ll know that your expected values will likely be safe estimates for money you can spend throughout the month. As an optional way to best utilize the created budget, I like to use a free phone app called *GoodBudget* that allows me to create categories of spending that I can record transactions against. I have a category created for each line item in this sub-section and as I spend throughout the pay period I record my transactions in the app, which subtracts that amount from the budget category. This makes it easy to ensure that I don’t go over budget in my spending money and that I’ll have enough in my checking account to cover the other cash outflows. There are plenty of free phone apps out there that exist for this purpose.

### Savings Analysis

The final sheet of your budget is the Savings Analysis page. This section is useful for tracking specific amounts you have saved for certain logical partitions or categories of savings, and for tracking spending against your savings.

#### Cash Inflows

The first section in this sheet requires no input from you, and that is the cash inflows section. This represents the amount of money by month that should be transferred from your checking account to your savings account(s) to cover the listed line items. The line items and numbers in this section are taken directly from the budget page, so as long as you have filled that out these items will be auto populated.

#### Cash Outflows

The “Cash Outflow” section represents any money that you spend from your savings account(s) in a particular month. If for example you have to buy new tires in a certain month, you would input the amount you need to spend in the monthly value cell and transfer that amount from your savings account to your checking to cover the cost.

The final section of this sheet is used to track the amount that you should have in your savings account at the end of a particular month based on the cash inflows/outflows to and from savings specified in the aforementioned sections. If the amount you have in your savings account(s) doesn’t match what is shown on this spreadsheet, chances are you have not transferred the amount that you should have from your checking account for that month. The only inputs here are the start values for each sub-section. If this is the first time using one of these templates you should fill in your current savings account balance for the “Actual Expected Balance” start value cell and then break up that value however you see fit in the logical categories, or even start at 0. If you are creating a new budget and have used these before, carry over the starting values from your previous year’s budget and input them in these cells. This will ensure that the values reflected on this sheet are always up to date, and that you are always aware of how much you have saved for various expected expenses, should the need arise.

## IV. Creating A New Budget To From An Existing One

If you’ve already created a budget using one of these templates in the past, and need to create a new one for the new year or for any other reason, you have two options: 

* You can either create a new budget manually using the steps above, with a couple additional steps at the end.

OR

* You can use one of the provided *CreateFromExisting* executor scripts to take care of this process for you. *(much less effort)*

### Manual Method

To create a new budget manually, simply follow the steps in the aforementioned section. Additionally, you should copy over any relevant cells from your existing budget to your new budget. The cells you may wish to copy over include:

* Any line items in column C of any of the sub-sections on the **Budget Analysis** page of your existing budget to the same cells on the budget of your new template copy.

* Any recurring values for the line items of the sub-section on the **Budget Analysis** page of your existing budget to the same cells on the budget of your new template copy. (if you would like to start from the same budget values)

* The “Expected Account Balances” December values from the **Savings Analysis** page (column S) of the existing budget to the “Expected Account Balances start values (column G). This is so that your new budget is aware of your current savings and logical savings category balances that have been tracked by your existing budget

### Automated Method

*NOTE:* you will need to have the Java runtime environment installed if you want to use this method.

If you would like this process taken care of with less steps invlolved, you can instead opt to use one if the provided scripts made for this purpose. If you are using windows, the file you’re interested in is the *CreateFromExisting.bat* file. If you’re using Linux, the file you’re interested in is the *CreateFromExsisting.sh* file. In either case there are two very simple steps:

* The first step is to open the file relevant to you in your text editor of choice and set the variable values at the very top of the file. Set EXISTING_BUDGET_PATH to the file path of your existing budget, and set NEW_BUDGET_PATH to the path where you would like your new budget to be created, including the name of your new budget. Keep in mind that for Windows users these paths should use double back slashes (\\\\) to separate directories, and for Linux users these paths should use forward slashes (/)

* The second step is simply to save your file, and then run by double clicking. On completion the newly created budget should exist at the path you specified with all of the aforementioned values filled in for you.

## V. Maintaining Your Budget Once Created

If you’ve gotten this far you will have successfully used one of the templates to create your budget, the hard part is over. From this point forward you will only need to edit your budget once per month. Specifically, you should take the following actions, which should be very quick, at the end of each month:

* On the **Budget Analysis** sheet update the actual monthly values for each cash inflow/outflow. Optionally recolor the cells to transparent to make it easier on your eyes next time you view your budget. If you have any “Excess Available Expenses” leftover at the end of these updates you will know you have some extra money to allocate towards savings, or spend on whatever you see fit.
* On the **Savings Analysis** sheet, use the “Actual Transferred In” (if you have two savings accounts like me) or “Total Transferred In” (if one account) monthly cell value to determine how much to transfer from your checking account to savings account(s) and make the transfer.
* On the **Savings Analysis** sheet input any money spent from savings in the “Cash Outflows” monthly cells and transfer the money to your checking account if not already done
* On the **Savings Analysis** sheet double check that the funds in your savings account(s) match the expected after all transfers are complete.  

[Optional] At the beginning of each month: if you use an app such as *GoodBudget* to track your spending, update the app according to the line items in the “Spending Money” sub-section on the **Budget Analysis** sheet.
