# Formulas in Excel

> ## What is a formula?
> This passage discusses how to analyze data using calculations in Microsoft Excel. It highlights the process of creating formulas, using the correct syntax, and understanding how Excel processes these calculations. The example provided involves updating prices and order amounts for Adventure Works, demonstrating how to create and modify a formula to calculate purchasing costs based on unit price and quantity. The passage also covers key concepts such as static and dynamic formulas, cell references, and editing formulas in Excel.
>
> ## Controlling calculations
> In this video, the importance of accurate and reliable calculations in Microsoft Excel is emphasized. It discusses how Excel processes calculations, introduces mathematical operators for addition, subtraction, multiplication, and division, and explains the concept of the order of precedence.
> 
> The example involving Jamie at Adventure Works demonstrates the need for correct calculation sequences, especially when combining multiplication and subtraction. The video highlights the use of parentheses to override the default order of precedence and control the sequence of calculations.
> 
> It also covers the two types of cell references in Excel: relative and absolute. Relative references adjust when a formula is copied to a new cell, while absolute references remain constant. The video explains how to make a cell reference absolute by adding a dollar sign.
> 
> Lastly, it mentions automatic recalculation in Excel and how to change the recalculation mode using Calculation Options.
> 
> Overall, the video provides valuable insights into Excel's calculation process and formula syntax for creating accurate spreadsheets.
>
> ## Controlling calculations in action
> In this video, the focus is on creating complex Excel formulas with multiple steps and controlling the syntax to ensure accurate calculations. Amy at Adventure Works is tasked with preparing a price quote for Contoso Bikes, including calculations for unit cost, discounts, and delivery charges.
> 
> The video explains how to build these calculations step by step:
> 1. Calculate the subtotal (cell G6): Multiply the cost per unit (E6) by the quantity ordered (F6).
> 2. Calculate the 10% discount (cell H6): Divide the subtotal (G6) by 100 and multiply by 10.
> 3. Calculate the total cost excluding delivery (cell I6): Subtract the discount (H6) from the subtotal (G6). This formula is also adjusted to include the multiplication of I2 for different outlets.
> 4. Calculate the total cost including delivery (cell J6): Add the total cost excluding delivery (I6) to the appropriate delivery charge based on the region (M2 for Region A and M3 for Region B). Parentheses are used to control the order of precedence.
> 
> The video emphasizes the importance of using parentheses to ensure the correct sequence of calculations and addresses the issue of operator precedence when subtraction and multiplication are involved.
> 
> Additionally, the video introduces the concept of absolute references, which are necessary when copying formulas using the autofill feature to prevent references from changing. It demonstrates how to make references absolute by adding dollar signs manually and by using the F4 key shortcut.
> 
> Finally, it shows how to use the autofill feature to copy the formulas to other cells efficiently, ensuring that the necessary references remain absolute.
> Overall, the video provides valuable insights into building and controlling complex Excel formulas for accurate calculations.


# Getting started with functions
> ## What is a function formula and how to write it
> In this video, the focus is on understanding and utilizing Excel function formulas to perform calculations. The example involves helping Lucas, an accountant at AdventureWorks, calculate the total quarterly sales for each regional sales team.
> The key points covered in the video are as follows:
> 1. **Introduction to Functions:**
   Functions are predefined formulas in Excel that perform calculations based on user-specified values. Functions help in executing more complex calculations and allow for dynamic content in response to changes in the worksheet.
> 2. **Function Categories:**
   Excel has many built-in functions categorized under various types such as financial, date and time, and math calculations, which can be accessed from the Formulas tab.
> 3. **Elements of a Function Formula:**
>    - **Function Name:** The name of the function (e.g., SUM).
>    - **Arguments:** Data or information provided to the function to perform calculations.
> 4. **Constructing a Function Formula:**
    A function formula begins with an equal sign followed by the function name, then the arguments enclosed within parentheses.
> 5. **Using the SUM Function:**
   An example is shown using the SUM function to calculate the total sales for a specific range of cells.
> 6. **Copying Formulas with Autofill:**
   The video demonstrates using the Autofill feature to copy a formula across cells, adjusting the references accordingly.
> 
> By the end of the video, the viewer should understand what functions are, be able to read the syntax of a function, and know how to use a function to perform calculations. The example involves calculating sales totals for regional sales teams at AdventureWorks using the SUM function.
>
> ## Using the INSERT function
> In this video, the focus is on using the Insert Function tool in Microsoft Excel to create function formulas accurately. The example involves helping AdventureWorks calculate annual sales totals for different regional teams using the SUM function.
> 
> Key points covered in the video are as follows:
> 
> 1. **Introduction to the Insert Function Tool:**
>    - The Insert Function tool in Excel helps users create function formulas with the correct syntax.
>    - It can be accessed from the Formulas tab or the worksheet screen.
> 2. **Using the Insert Function Tool:**
>    - Position the cursor in the cell where the function result should appear.
>    - Access the Insert Function tool using one of two methods: through the Formulas tab or the worksheet screen.
>    - The Insert Function dialog box opens, providing access to various functions.
> 3. **Function Categories:**
>    - Functions are grouped into categories (e.g., Math and Trigonometry) for easier navigation.
>    - Users can select a category to access a specific list of functions.
> 
> 4. **Selecting a Function:**
>    - Users can choose a function from the list or use the search bar to find a specific function.
>    - Hovering over a function name displays a short description, and clicking "Help on this function" opens the Microsoft support page for more information.
> 5. **Function Arguments Dialog Box:**
>    - After selecting a function, the Function Arguments dialog box appears.
>    - It displays the required and optional arguments for the selected function.
>    - Excel may suggest a cell range based on the context.
> 6. **Navigating and Editing Selection:**
>    - Users can use the navigation buttons to change the selected cell range or edit the formula as needed.
> 7. **Formula Result and Warnings:**
>    - The dialog box displays the formula result, which can be a total or a value depending on the function.
>    - Warning messages may appear if there are errors in complex function formulas.
> 8. **Completing the Function Formula:**
>    - After confirming the function and arguments, users can select "OK" to insert the formula into the worksheet.
> 
> In the example, the SUM function is used to calculate annual sales totals for Team A at AdventureWorks. The Insert Function tool helps ensure that the formula syntax is correct and accurately targets the required data. Once the formula is created for Team A, it can be copied to generate sales totals for other teams.
> 
> By using the Insert Function tool, viewers should become familiar with its categories, understand how to create function formulas, and recognize its utility in ensuring accurate calculations.