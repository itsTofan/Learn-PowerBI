# Using functions to clean or standardise text
> ## How inconsistent data can affect data analysis
> In this video, the importance of accurate and trustworthy data for data analysis is highlighted. Common data errors in Microsoft Excel are discussed, along with their potential negative impact on data analysis. These errors include:
> 1. Misspelled Names or Key Phrases: Misspelling can lead to Excel not recognizing entries or linking them to other important details correctly.
> 2. Unnecessary Characters: Entries containing unnecessary characters, such as a dollar sign before numbers, can be treated as text, causing issues in calculations.
> 3. Unnecessary Spaces: Extra spaces before or after entries can affect data analysis as Excel considers them different from entries without spaces.
> 4. Misplaced Entries: Putting an entry in the wrong column or under an incorrect heading can lead to misclassification during analysis.
> 5. Inconsistent Layout or Content: Inconsistent data presentation or layout can create difficulties in data processing and analysis.
> 6. Abbreviations and Acronyms: Using inconsistent abbreviations or acronyms can lead to confusion during data analysis. Standardizing abbreviations is recommended.
> 7. Date Format: Dates should be entered in a specific format (e.g., month/day/year) to ensure Excel recognizes them as dates for time-based analysis.
> 8. Duplicate Data: Duplicate entries in a dataset can distort analysis results by counting items multiple times.
> 
> The video emphasizes the importance of data quality and offers tips for avoiding common data errors, such as designing efficient spreadsheets, standardizing data entry, and checking for duplicate data. Overall, the goal is to ensure reliable and error-free data for accurate data analysis.
>
> ## Date and time functions
> In this video, the importance of date and time calculations in Excel for businesses is highlighted. Date and time calculations are essential for planning, tracking performance, and analyzing data. Here are some key points from the video:
> 
> 1. **Planning for Increased Demands:** Date and time calculations help businesses plan for increased demands on products and resources. For example, a construction project manager can use Excel to calculate the number of working days between project start and end dates.
> 2. **Dynamic Planning:** Excel allows for dynamic planning by updating calculations as time passes. This helps in adjusting schedules and plans as needed.
> 3. **Identifying Performance Trends:** By organizing results by date, businesses can identify patterns and trends in performance. This can help in understanding the factors that influence changes in performance, whether internal or external.
> 4. **Tracking Transactions:** Business transactions are often recorded with dates and times. Excel's date and time functions are useful for tracking and analyzing these transactions.
> 5. **Serial Numbers:** Excel uses serial numbers to track calendar days. Each date is assigned a serial number starting from January 1, 1900. These serial numbers are used in calculations.
> 6. **Date Functions:** Excel provides various date functions, such as TODAY and NOW, for displaying the current date and time. Functions like MONTH, DAY, and YEAR allow you to extract specific components of a date.
> 7. **DATE Function:** The DATE function allows you to create dates by specifying the year, month, and day as arguments.
> 
> By understanding and using these date and time functions and formulas, businesses can efficiently plan, track, and analyze their data, leading to better decision-making and insights.
> 
> Jamie at Adventure Works Distribution Hub can use these functions to track delivery times, dates, and intervals for their purchases and customer orders. Understanding these concepts will enable her to work more effectively with date and time-based data.
>
> ## Dynamic date and time entries
> In this video, you learned how to work with date and time functions and formulas in Excel to create dynamic date and time entries for various purposes. Here are the key points from the video:
> 
> 1. **Serial Numbers:** Excel uses serial numbers to represent dates. Each date is assigned a unique serial number, starting from January 1, 1900. This serial number is used for date calculations.
> 2. **Calculating Date Differences:** You can calculate the number of days between two dates using a simple subtraction formula. For example, subtract the start date from the end date to find the number of project days.
> 3. **Dynamic Date Entry:** You can use the `TODAY()` function to display the current date in a cell. This date updates automatically every 24 hours, ensuring it always reflects the current date.
> 4. **Calculating Days to a Future Date:** To calculate the number of days remaining until a future date, subtract the future date from the current date (using the `TODAY()` function). This formula updates daily.
> 5. **Extracting Year from a Date:** You can use the `YEAR()` function to extract the year component from a date. This is useful when you need to work with the year separately from the full date.
> 6. **Copying Formulas:** To apply date and time formulas to multiple cells, you can use the AutoFill feature. Simply double-click the small square in the bottom-right corner of the cell with the formula, and Excel will copy the formula down the column.
> 
> By applying these techniques, you can create dynamic date and time entries, calculate date differences, and extract specific date components, making your Excel worksheets more versatile and efficient.
> 
> Adventure Works used these formulas to track campaign launch dates, project timelines, and year components for various countries, allowing for better project management and planning.
>
> ## Logical functions
> In this video, you learned about logical functions in Microsoft Excel, focusing on the IF function. Logical functions are used to make decisions in Excel based on certain conditions or criteria. The IF function allows you to perform actions depending on whether a specified condition is met.
> 
> Key points covered in the video:
> 1. Logical functions are used to ask yes or no questions about data in Excel.
> 2. Logical operators, such as equal to, greater than, less than, greater than or equal to, less than or equal to, and not equal to, are used to compare values in logical functions.
> 3. The IF function in Excel requires three arguments:
>    - Logical_test: The condition or test to be performed on a cell's value.
>    - Value_if_true: The action to take if the condition is met (returns TRUE).
>    - Value_if_false: The action to take if the condition is not met (returns FALSE).
> 4. Logical functions help automate decision-making processes based on data.
> 5. You can use the IF function to calculate bonuses, make grading systems, and perform various conditional calculations in Excel.
> 
> Overall, the video demonstrated how to use the IF function to create logical formulas in Excel, making it possible to automate decisions and actions based on specific criteria or conditions in your data.
>
> ## Using nested IF and IFS function
> In this video, you learned about nested IF and IFS functions in Microsoft Excel. These functions allow you to test for multiple conditions and perform actions based on the results of those conditions.
> 
> Key points covered in the video:
> 
> 1. Nested IF functions: You can nest multiple IF functions inside each other to perform a series of elimination tests. Each IF function checks a specific condition, and if the condition is met (returns TRUE), it performs an action. If the condition is not met (returns FALSE), it moves on to the next IF function.
> 2. IFS function: The IFS function is designed to simplify the process of testing multiple conditions. It allows you to specify a series of conditions and their corresponding actions in a single formula. Excel checks each condition in order, and when it finds a condition that is met, it performs the associated action and stops evaluating the remaining conditions.
> 3. Syntax of nested IF function: The syntax of a nested IF function includes the logical test, value if true, and value if false for each condition. You can nest multiple IF functions by using them as the value if true or value if false argument in another IF function.
> 4. Syntax of IFS function: The IFS function begins with the IFS keyword, followed by pairs of logical tests and corresponding values if true. There is no need for a value if false argument in the IFS function.
> 5. Nested functions can help automate decision-making processes based on multiple criteria in your data.
> 
> Overall, the video demonstrated how to use nested IF and IFS functions to handle complex conditional logic in Excel, making it possible to perform different actions based on multiple conditions in your datasets.
>
> ## AND and OR functions
> In Microsoft Excel, you can use the AND and OR functions to perform multiple logical tests simultaneously and make decisions based on the results. Here's a summary of these functions:
> 
> **1. AND Function:**
> - The AND function allows you to perform multiple logical tests (up to 30 arguments) simultaneously.
> - Each logical test generates a result of either TRUE or FALSE.
> - For the AND function to return TRUE, all the individual logical tests must also return TRUE. If any one of them returns FALSE, the AND function returns FALSE.
> - Syntax: `=AND(logical_test1, logical_test2, ...)`
> 
> **2. OR Function:**
> - Similar to the AND function, the OR function allows you to perform multiple logical tests (up to 30 arguments) simultaneously.
> - The OR function returns TRUE if at least one of the logical tests is TRUE. It returns FALSE only if all logical tests are FALSE.
> - Syntax: `=OR(logical_test1, logical_test2, ...)`
> 
> **Using AND and OR Functions:**
> - You can use these functions to test conditions in your data and make decisions based on the results.
> - They can be especially useful when combined with the IF function. For example, you can nest an AND function within an IF function to test for multiple criteria simultaneously.
> - Example: `=IF(AND(condition1, condition2), "Result if true", "Result if false")`
> 
> **AND Function Example:**
> - If you want to check if three values (A2, B2, and C2) are all less than 200, you can use `=AND(A2<200, B2<200, C2<200)`. This formula returns TRUE only if all three conditions are met.
> 
> **OR Function Example:**
> - If you want to check if any of three values (A2, B2, or C2) is less than 200, you can use `=OR(A2<200, B2<200, C2<200)`. This formula returns TRUE if at least one condition is met.
> 
> By using the AND and OR functions, you can perform complex conditional tests in Excel and make data-driven decisions based on multiple criteria simultaneously. These functions are valuable tools for data analysis and decision-making within spreadsheets.
>
> ## IF functions
> In summary, the reading covers three useful functions in Excel that can be incorporated with the IF function to enhance data analysis and decision-making:
> 
> **1. SUMIF:**
> - Use the SUMIF function to calculate the sum of values in a range based on a specified condition.
> - Syntax: `=SUMIF(range, criteria, sum_range)`
> - Example: `=SUMIF(B2:B24, "seattle", G2:G24)` sums the values in column G where the corresponding cell in column B is "Seattle".
> **2. AVERAGEIF:**
> - The AVERAGEIF function calculates the average of values in a range that meet a specified condition.
> - Syntax: `=AVERAGEIF(range, criteria, average_range)`
> - Example: `=AVERAGEIF(B2:B24, "seattle", G2:G24)` calculates the average of values in column G for rows where the corresponding cell in column B is "Seattle".
> **3. COUNTIF:**
> - COUNTIF function counts the number of cells in a range that meet a given condition.
> - Syntax: `=COUNTIF(range, criteria)`
> - Example: `=COUNTIF(B2:B24, "seattle")` counts the occurrences of "Seattle" in column B (B2 to B24).
> 
> These functions help to summarize and analyze data based on specific conditions, providing a more dynamic approach to data manipulation and analysis in Excel. By using logical functions like IF in combination with these specialized functions, you can perform intricate data operations and make informed decisions based on your data.