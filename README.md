java c
ECON10151   Lecture   6 
Introduction to Visual   Basic for Applications   (VBA) 
Nov   2024
Learning Outcomes 
•    Be able to implement custom VBA functions within   Excel   to   perform   complex   calculations
•    Be able to   understand and   apply VBA   loop   structures,   allowing   them   to   automate   repetitive   calculations   across datasets
Introduction This lecture introduces students to the power of   Excel VBA as   a tool for   automating   calculations   and   efﬁciently   handling   repetitive tasks   in   data   processing.   We will   start with   creating   a   custom   VBA function,   which   allows   us   to   perform   complex   calculations   that   Excel’s   built-in   functions   might   not   support   directly.    Building   on   this,   we’ll explore   how to use VBA   loops to extend these functions   across   multiple   rows,   automating the   calculation   for an entire dataset in just a few steps.   Writing the code yourself is not a requirement   for this course;   however,   it   is   more   important to   understand the code and   its   meaning.
1 Set Up 
The Visual   Basic editor   is   located   under the   Developer tab.   However, the   Developer tab   is   not displayed   in the   Excel   ribbon   by default,   but you can easily add   it   to   access   VBA   editor.
•    For Windows   Users
1.    Click   on the   File   tab.
2.    Select Options.
3.    In the   Excel Options window, choose   Customise   Ribbon.
4.    In the Customise the   Ribbon section, locate the Main Tabs   list   and check the   box   next to   Developer.
5.    Click OK to   apply the   changes.
• For   Mac   Users
1.    Click on   Excel   in   the   top   menu.
2.    Preferences.
3.    Choose   Ribbon    Toolbar.
4.      the Customise the   Ribbon section, ﬁnd the   Main Tabs   list and   check the   box   next to   Developer.
5.    Click Save to conﬁrm   the   changes.
2 Introduction to Visual Basic for Applications (VBA) Visual      Basic   for   Applications      (VBA)      is   a   programming      language   developed   by      Microsoft   that      is   embedded   within      Excel   and   other   Ofﬁce   applications.      VBA   enables   users   to   automate   repetitive   tasks,   enhance   data   analysis, and create custom functions and solutions tailored to speciﬁcneeds.
Overview of VBA Editor 
The VBA   Editor   is the environment within   Excel where you write   and   edit   VBA   code.
•    Code Window: Where you type and   edit   your   VBA   code.
•    Project    Explorer:    A    navigation   pane   that   shows   all   the   open   VBA   projects   and   the   components,      like   modules and worksheets, within them.
•    Properties Window: Allows you to view and   modify   properties for   selected   objects.
•    Immediate Window:    Useful for testing   and debugging code   on   the   spot   by   allowing   you   to   run   individual   lines or   commands.
You can access the VBA   Editor by   going to the   Developer   tab   in   Excel   and   selecting   Visual   Basic.   In the VBA code window, different colours are   used to   represent   various   elements   of   the   code.
• Black:
Standard Code:   This is the default colour for regular text,   including variable   names, function   names,   and   most code statements.
• Blue:
Keywords:   This colour is used for VBA keywords   and   reserved words   (e.g.,   Sub,   Function,   End,   If, Then,   Else,   For,   Next).   Keywords are commands that VBA   recognises and   interprets.
• Green:
Comments:   Any text following an apostrophe   (’) appears   in   green.   Comments   are   not   executed   and   are   used to explain the code or   provide   notes for   the   programmer.
Please ﬁnd   below a screenshot of the VBA   Editor window   along with   an   illustration.

3 Practice with VBA 
Task 
In country A, the amount of   income tax   owed depends   on the   individual’s   annual   income,   as   follows:
Table   1:   Tax   Rates and   Bands
Tax Rate 
Income 
0% 
Up to £13,000 
20% 
£13,000.01 to   £50,000 
40% 
Above   £50,000 
In the   Excel ﬁle called   L6 data,   it contains the annual   income for   ten   individuals.
Please calculate how much tax each individual should pay. 
Understanding of This Task 
1.      If x ≤   13;   000:
The tax   rate   is 0%.   The person   has   to   pay:
Tax   =   0
2.      If   13;   000   < x ≤ 50; 000:
The tax   rate   is 20% for the amount above £13,000.   The   person   has   to   pay:   Tax   =   0.2   × (x   —   13;   000)
3.      If   x   ≥   50;   000:
• The tax   rate   is 20% for the amount   greater than £13,000 and   less   than   or   equal   to   £50,000.
• The tax   rate   is 40% for the amount greater   than   £50,000.
The person   has to   pay:
Tax   =   0.2   × (50;   000 — 13;   000)   + 0.4   ×   (x —   50;   000)
Approach 1: Create a VBA Function for Income Tax Calculation 
1.    Open the   VBA   Editor
•    Open your Excel workbook, go to the   Developer tab, and   click   on   Visual   Basic.
(Altenatively,   Press Alt   +   F11 to open the Visual   Basic for   Applications   (VBA)   editor.)
2.    Insert   a   New   Module
•    In the VBA editor,   go   to   the   menu   and   click   on   Insert.
•    Select Module.   This will create a   new   module where you can write   your   code.
3.      Write the   Function
To create a   new function,   please   read through the   following   procedure,   including the   code   and   explana-   tion.   Once you   understand them, enter the following VBA code   into the   new   module:
• Deﬁne the Function: 
Function CalculateIncomeTax(income As   Double) As   DoubleThis   deﬁnes   a   function      named   CalculateIncomeTax   that   takes   a   single   arg代 写ECON10151 Lecture 6 Introduction to Visual Basic for Applications (VBA)Matlab
代做程序编程语言ument      (income)   and   returns   a Double.   Double is   short for   ”double   precision ﬂoating-point”   and   is   typically   used   to   store   numbers that   require a   large   range and/or precision,   including decimals.
• Variable Declaration:
Dim Tax   As   Double
This declares a variable to   hold the calculated tax.
• Specify If Statements:
If income <=13000 Then
Tax = 0
ElseIf income > 13000 And income <=50000 Then
Tax = 0.2 * (income - 13000)
Else
Tax = 0.2 * (50000 - 13000) + 0.4 * (income - 50000)
End If

The ﬁrst   If checks if the income is less   than   or   equal   to   £13,000   and   sets   the   tax   to   0.   The   second   ElseIf   checks   if   the   income   is   between   £13,000.01   and   £50,000   and   calculates   the   tax   at   20%.   The ﬁnal Else block calculates the tax for   incomes above £50,000   using   the   respective   rates.
• Return Value:
CalculateIncomeTax = tax
This   assigns the   calculated tax to the   function   name,   which   returns   the   value   when   the   function   is   called.
• Function Completed:
End Function
This   indicates that the deﬁnition of the function   has been   completed.   The code that you should type   into the   module will   be   displayed   below.



4.    Save Your   Work
•    Make sure to save your   Excel   workbook   as   a   Macro-Enabled   Workbook   (.xlsm   format)   to   preserve   the VBA   code.
5.    Use the   Function   in   Excel
•    In any cell in your   Excel worksheet,   you   can   now   use   your   new   function.
For example,   if you want to calculate the tax for   the   ﬁrst   person,   you would   type   the   following   in   cell   C2:
=CalculateIncomeTax(B2)
Press   Enter, and   Excel will return   the calculated income tax based on the   rules deﬁned   in your VBA   function.
•    Now, you can apply the function to   determine the   tax   that   should   be   paid   by the   other   individuals   in   the dataset.
Approach 2: Using Loops for Batch Income Tax Calculation Alternatively,   we   can   use   a      loop   in   VBA   to   calculate   the   income   tax   for      multiple   individuals   at   once.         This   method allows you to process each individual’s income in a single   procedure.   Follow these steps to   implement   this approach:
1.    Open   VBA   editor
•    Go to the Developer tab   in the   Excel   ribbon.   Click on Visual   Basic   to   open   the   VBA   editor.
•    Click on Module   1.   We will write   new code   in the section   below where we   created   the   function.
2.      Write the   Loop   CodeTo code the   procedure, we   need to clarify what we want   to   achieve.   We   will   use   the   range   C2   to   display   the tax   amounts   owed   by   each   individual.   The ﬁrst   output cell   is   C2.   To calculate the   tax   for   C2,   we   will   use the      CalculateIncomeTax    function that we just created.   The input for this function will be the value in   cell   B2.   Excel will pass the value of   B2 to the    CalculateIncomeTax      function   and   return   the   result   in   C2.Next,   Excel   will   move   to   the   following   row,   using   the   value   in   B3   as   input   for   the   function   and   returning   the   tax   amount   in   C3.    This   process   will   continue   for   each   row   until   we   reach   the   last   individual   in   the   dataset.
Let’s follow these steps to code   the   procedure:
• Deﬁne the Procedure:
Sub CalculateTaxesForAll()
• Deﬁne Variables: 
Dim   i As   Integer
This   line declares a variable   i as an   integer, which will be   used   as   a   loop   counter.
Dim   income As   Double
This   line declares a variable income as a double, which will   store   the   income   value for   each   individ-   ual as a ﬂoating-point   number   (allowing for   decimals).
Dim lastRow As Long
This   line declares a variable   lastRow as a   long   integer, which will   hold the   number   of the   last   row   in   the   income column with data.
• Specify the number of rows in the dataset
lastRow = Cells(Rows.Count, ”B”).End(xlUp).RowThis   line   ﬁnds   the   last   row   of   data   in   column   B(income   values).         It   starts   from   the   bottom   of   the   worksheet   (Rows.Count   gives   the   total   number   of   rows)   and   moves   up   until   it   ﬁnds   the   last   non-   empty cell   in column   B. This value   is stored   in   the   lastRow variable.
• Deﬁne the procedure:
For i = 2 To lastRow
This line initiates a   For loop that starts at   row 2   (the ﬁrst   row of   data)   and continues to   lastRow.   The   loop   iterates through each   row containing   income data.
income = Cells(i, 2).Value
This   line   reads the   income value from column   B   (the ﬁrst column) of the current   row   i and   assigns   it   to the   income variable.
Cells(i, 3).Value = CalculateIncomeTax(income)
This   line   calls   the   previously   deﬁned CalculateIncomeTax function,   passing   the   income   value   as
an   argument.      It takes the   result   (the calculated tax)   and writes   it   to   column   C   (the   third   column)   of   the same   row   i.
Next i
This   line   indicates the   end   of the   loop.    The   loop   continues with the   next value   of   i   until   it   reaches   lastRow.
The code that you should type   into the   module will   be   displayed   below.

3.    Run the Code   and   Review   the   Results
•    Click on the worksheet TaxCalculation2, then go   to   the   Developer   Tab.
•    From the ribbon   menu, click on   Macros, select CalculateTaxesForAll,   and   then   click   on   Run.
You will   now see the calculated taxes for   all   individuals   in the   dataset.

         
加QQ：99515681  WX：codinghelp  Email: 99515681@qq.com
