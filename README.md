java c
Lecture   1:   Basic   Data   Analysis   using   Excel
ECON10151:   Computing   for   Social   Scientists
September   22, 2024In this lecture, we’ll introduce some key features of   Excel that are especially useful when working with   social   science   data.   Excel has a variety of   tools designed to make data analysis quicker and more efficient.   Today, we’ll focus on essential functions   and   techniques   that   will   help   you   clean   and   format   datasets,   create   new   variables,   carry   out basic   statistical   calculations,   and   generate   simple   visualisations.    By   working   through   practical   examples,   you’ll   gain   hands-on   experience,   equipping   you   to   apply these tools to your own data with   confidence.
1            Data
The   dataset   we   will   be   working   with   contains   the   average   price   and   quantity   of hand   washing   products   sold   in   the   UK   each   month during 2020.   This data was sourced from the Kantar FMCG   Purchase   Panel.
For this exercise, we will   assume:
•    The   firm’s   variable   cost   is   £3 per   unit   of   quantity.
Variable costs are expenses that fluctuate with production or   sales volume.   Common   examples   include:
–    Raw materials:   The cost of materials required to produce each unit.
–    Direct labour:   Wages paid to workers based on the number of units produced or hours worked.
–    Packaging:   Costs associated with packaging each product.
•    The   firm’s   fixed   cost   is   £8,500.
Fixed costs remain the same, regardless of   production levels.   Typical fixed costs   include:
–    Rent   payments:   The   cost   of   leasing   a   building   or   office   space.
–    Insurance:   Premiums paid for coverage, which do not vary with output.
–    Depreciation:   The gradual reduction in   the value of fixed assets like machinery or   equipment.
Please download the spreadsheet titled ‘raw data’ from Blackboard to   get   started.
2          From   Raw   Data   to   Profit   Analysis
We’ll follow a step-by-step approach to analyse   the profit   from   sales.
2.1          Cleaning   DataProperly organising your data is the first crucial step in preparing it for analysis.   Sometimes, we deal with imported   or unstruc-   tured   data   where   multiple   pieces   of   information   are   combined   into   a   single   cell, which   needs   to   be   split   for   proper   analysis—just like the raw data we’re working   with.
In Excel, the   Text   to   Columns function is   an effective   data-splitting   tool   for   separating   data   contained   in   one   column   into   multiple columns, based on a specific delimiter   (such as   a   comma,   space,   or   tab).   Here’s   how   to   use   it:


Step   1    Select the Data:   Highlight the cells containing the text you want   to   split.
Step   2    Open   ‘Text   to   Columns’:   Go   to   the   Data   tab   and   select   Text   to   Columns.
Step   3    Choose   ‘Delimited’   :   In   the   dialog   box, choose   Delimited, as   our   data   is   separated   by   commas.
Step 4    Set Delimiter:   Check the   Comma option as the delimiter (since   the   values   in   the   data   are   comma-separated).
Step 5    Finish:   Excel will show a preview of how the   data will   look.   Once   you’re   happy   with   the   result,   click   Finish.After   using   Text   to   Columns   to   split   the   data, you   might   notice   that   some   of   the   headers   are   spread   across   multiple   columns.      In   such   cases, manually   adjust   the   headers   by   combining   the   text   and   labelling   them   appropriately   (e.g., “Quantity   Sold   (Litres)”   or   “Price per Litre”).    Ensure that   each row   of data   aligns   correctly   with its   respective   column   before   proceeding   with   further      analysis.
2.2          Transposing   Data
Sometimes your data is organised horizontally,   but you need it   vertically—or   vice   versa.   For   example,   your headers   may   be   in   a row when you need them in a column, or data   may be   listed   in   columns but   would be   more   useful   in rows.
Transposing data in Excel means flipping the   orientation,   converting rows   into   columns   or   columns   into   rows.   This   can   be   very useful when the data’s current format doesn’t suit the   analysis   or   visualisation   you’re   aiming   for.
We’ll explore two methods to transpose data in Excel during this lecture.
1.    Paste   Special   Transpose:
Step   1    Select the Data:   Highlight the range of cells you want to transpose, including headers if necessary.
Step   2      Copy   the   Data.
Step   3    Choose   the   Destination:    Click   the   cell,   e.g.   A7,   where   you   want   the   transposed   data   to   appear.    Ensure   there’s   enough space, as transposing will expand the data either   vertically   or horizontally.
Step   4    Paste   Special:   Navigate   to   Home   tab   →   Paste   →   Paste   Special, or   right-click   on   the   destination   cell   and   choose Paste   Special.
Step   5    Transpose   the   Data:    In   the      Paste      Special   dialog   box,   select   Values   (to   paste   only   the   values)   and   check   the Transpose option at the bottom.   Then click OK. Your data   will   now   be   flipped   between   rows   and   columns.
2.    TRANSPOSE Function:
The   TRANSPOSE   function   is   another   way   to   transpose   data,   and   it   differs   from   the   Paste   Special   method   in   that   it creates a dynamic link.   This means that if the original data changes, the transposed   data will   update   automatically.
The   syntax is:
= TRANSPOSE(array)
Here, the array represents the range of cells to be transposed.
Follow these   steps:
Step   1    Choose the Destination:   Again, make sure there’s enough   space.
Step 2    Enter the Formula:   In the selected range, type the   formula
= TRANSPOSE(A1 :   N4)
                       where   A1 :   N4 is   the   range   of   your   original   data.
Step   3      Press   Enter.
Note: If   you’re   using   older   versions   of   Excel,   press   ‘Ctrl   +   Shift   +   Enter’ (Windows)   or   ‘Command   +   Return   ’   (Mac), as the   TRANSPOSEfunction requires this to work as an   array   formula.


- Summary:
•    Use   TRANSPOSE if you need   a   live   link between your   original   and   transposed   data.    This   is   ideal   for   datasets   that   are   regularly updated.
•    Use   Paste   Special when you need a one-time,   static   rearrangement   of your   data   and   don’t require   it   to   update   automati-   cally.
2.3          Removing   Duplicates
When   importing   data   from   external   sources   such   as   CSV   files, databases, or   surveys, it’s   common   to   encounter   duplicate   entries.   Redundant data can skew your analysis, so it’s   essential   to remove   any   duplicate   values before proceeding.Manually   identifying   and   removing   duplicate   rows   can   be   time-consuming,   especially   with   large   datasets.      Fortunately,   Excel’s    Remove    Duplicates   feature   helps   you   quickly    find   and   eliminate   duplicate   entries,   ensuring   your   dataset   remains   accurate and free of   redundancy.
Follow these steps to remove duplicates in Excel:
Step   1    Select the Data Range:   Highlight the range of cells where you want to remove duplicates.   This can be a single column   or an entire table, depending on   what   you   need.
Step   2    Open   the   ‘Remove   Duplicates’ Tool:   Go   to   the   Data   tab   and   click   on   Remove   Duplicates.
Step 3    Choose   Columns   to   Check for   Duplicates:    A   dialog box   will   appear   showing   all the   columns   in   the   selected   range.   Tick or untick the boxes to specify which columns should be   checked   for   duplicates.
•    If you   select just   one   column,   Excel   will   remove   rows   where   the   value   in   that   column   is   repeated,   even   if other   values   in   the   row   differ.
•    If   you   select   multiple   columns,   Excel   will   treat   rows   as   duplicates   only   if   the   combination   of   values   in   those   columns is identical.
Step 4    Click   OK:   Once   you’ve   chosen   the   columns   to   check,   click   OK.   Excel   will   display   a   message   showing   how   many   duplicates were found and removed, and how many unique entries remain.
- Important Notes:
•    Removing duplicates is permanent.   Consider creating a backup of   your dataset before using this tool to prevent accidental   data loss.
•    The   Remove   Duplicates   tool   is   case-insensitive, meaning   it   treats   ‘Judith’ and   ‘judith’   as   duplicates.
2.4          Creating   New   VariablesOften, creating new variables is essential for deeper analysis.   In this section, we will   create   three   new   variables:   Total   Revenue   (TR), Total Cost (TC), and Profit.   Additionally,   we will   categorise   the   data   by   creating   a   dummy   variable   and   a   string   variable   to indicate whether a profit was   made.
1.    Calculating   Total   Revenue, Total   Cost,   and   Profit
We will start by calculating Total Revenue (TR), Total   Cost   (TC),   and   Profit   using   simple   Excel   formulas.   Step   1    Total Revenue (TR): This represents the total income from   sales   and   is   calculated   as:
TR   = Quantity   Sold   ×   Price
In   Excel, you   can   enter   this   formula   in   a   new   column.   For   example, in   cell   E8,   type:
= C8   * D8
where   C8 contains   the   quantity   sold, and   D8 contains   the   price   for   January.   Step   2    Total   C代 写ECON10151: Computing for Social Scientists Lecture 1: Basic Data Analysis using ExcelPython
代做程序编程语言ost   (TC): Total   Cost   is   the   sum   of   variable   and   fixed   costs:
TC = (Variable Cost   ×Quantity Sold)+Fixed Cost
Recall   that   the   variable   cost   is   £3   per   unit,   and   the   fixed   cost   is   £8,500.   Enter   these   values   in   cells   B21   and   B22   respectively.
In   Excel, you   can   enter   the   formula   as   follows   in   cell   F8:
= (3*C8)+$B$22
Here, $B$22   locks   both   the   column   and   row   reference   to   cell   B22, which   contains   the   fixed   cost   (£8,500), ensuring that the reference cell doesn’t change when copying the formula down to   other rows.
Step 3    Profit:   Profit is calculated as the difference between   Total   Revenue   and   Total   Cost:
Profit = TR−TC
In   Excel, type   the   formula   in   cell   G8 as:
=   E8   −   F8
where   E8 contains   TR, and   F8 contains   TC.
2.    Categorising Profit with a Dummy   VariableNext,   we   will   categorise   profit   by   creating   a   dummy   variable.    A   dummy   variable   is   a   numeric   variable   that   takes   the   value   1 or   0.   Here, we   shall   use   1 to   represent   a   positive   profit,   and   0   will   represent   no   profit   or   a   loss.   We’ll   use   Excel’s   IF function to create this variable.
The   IF function performs a conditional test and returns one value if   the condition is TRUE and another if   the condition is   FALSE. The syntax is   as   follows:
= IF(logical_test,          value      if      true,          value      if      false)
where:
•    logical_test:   The   condition   to   check   (e.g.,   G8 > 0 to   test   if   profit   is   positive).
•    value_if_true:   The   value   returned   if   the   condition   is   TRUE   (e.g.,   1).
•    value_if_false:   The   value   returned   if   the   condition   is   FALSE   (e.g., 0).
In   this   case, to   create   the   dummy   variable, type   the   following   formula   in   a   new   column   (e.g.,   cell   H8):   = IF(G8   > 0,   1,   0)
This   formula   assigns   a   value   of   1 if   the   profit   (in   G8)   is   positive, and   0   if   it   is   zero   or   negative.
3.    Creating   a   String   Variable   for   Profit
To make the data more descriptive, we can create a string variable that labels whether a profit was made.   Instead of   using   numeric   values, we   will   use   "YES" for   profit   and   "NO" for   no   profit.
In   a   new   column   (e.g.,   I8), use   the   following   formula:
= IF(G8 > 0, "YES","NO")
This   formula   will   display   "YES" if   the   profit   is   positive   and   "NO" if   it   is   not.
Note:   When   using   text values   in Excel   formulas, always enclose   them   in quotation   marks   ("       ").   If you don’t,   Excel   will   return   an   error   (#NAME?).
2.4.1          Applying   Formulas   Across   Rows   Using   the   Fill   Handle
Once you’ve entered your formulas, you can quickly apply them   to   the   entire   dataset   using   the   Fill   Handle.   To   do   this:
•    Select   the   cell(s) containing   the   formula(s) you   want   to   copy.
•    Hover   over   the   bottom-right   corner   of   the   selected   cell(s) until   a   small   square   appears—this   is   the   Fill   Handle.
•    Click and drag the   Fill   Handle down or across   to   fill   the   formula   into   the   adjacent   cells.
This method allows you to efficiently apply the formulas to   all relevant rows   or   columns.
2.5          Conditional   Formatting
Conditional   formatting   is   a   powerful   tool   in   Excel   that   helps   you   visually   emphasise   specific   values   in   your   dataset.    In   this   exercise, we’ll use it to highlight the cells   that display   "YES" or   "NO" for profit,making patterns   easier to   identify   at   a   glance.
Step   1    Select   the   Range   of   Cells:   Highlight   the   cells   where   you   want   to   apply   conditional   formatting,   e.g.   I8:I19.   Step   2    Open   the   Conditional   Formatting   Menu:   Navigate   to   Home   →   Conditional   Formatting.
Step   3    Create   a   Rule   for   "YES":
Select   New   Rule... from   the   drop-down   menu, then   choose   the   Classic   style.   Next, select   Only   format   cells   that   contain from   the   list   of   rule   types.   In   the   options, choose   Specific   Text,   select   containing,   and   type   YES   in   the   text   box.
Note:   You can choose   from a variety of   pre-defined rules or create your own custom   rule   for more specific needs.
Step   4    Choose   a   Format:
In   the   Format   with   drop-down,   select   a   fill   colour   to   highlight   the   cells   containing   "YES"   (e.g.,   Green      Fill   with      Dark Green   Text).   Click   OK   to   apply   the   rule.
Step   5    Create   a   Rule   for   "NO":
Repeat   the   process   for   the   cells   that   contain   "NO", but   choose   a   different   fill   colour   (e.g.,   Red   Fill   with   Dark   Red   Text)   to distinguish them from the   "YES" entries.
2.6          Data   Analysis   Functions
To wrap up,   we’ll use   some   of Excel’s built-in functions   to   summarise   and   analyse   our   data.    These functions   will   allow   us   to   quickly count, calculate, and identify key metrics related to our profit data.
Step   1    Counting   Profitable   Months:
To   count   the   number   of months   where   profit   was   made,   we   can   either   use   the   COUNTIF   function   or   sum   the   Profit   dummy variable.
•    Using   COUNTIF: This function counts the   number   of cells   that   meet   a   specific   condition.   To   count   the   months   with profit, use the following formula:
= COUNTIF(I8 :   I19,   "YES")
This   counts   the   number   of   cells   in   the   range   I8:I19 that   contain   "YES".
•    Using   SUM:   If you’ve   created   a   dummy   variable   for   profit   (1   for   profit,   0   for   no   profit),   you   can   use   the   SUM   function to count the number of   profitable months:
= SUM(H8 :   H19)
This   adds   up   the   values   in   the   dummy   variable   column   (H8:H19),   effectively   counting   how   many   months   were profitable.
Step   2    Calculating   the   Average   Profit:
To find the average profit, use the AVERAGE function.   This will   calculate the   mean   of the profit   values   in   your   dataset.   Enter the following formula in a new   cell:
= AVERAGE(G8 :   G19)
This   formula   calculates   the   average   profit   from   the   values   in   the   range   G8:G19.
Step 3    Finding the Minimum and Maximum   Profit:
You can quickly identify the lowest and highest profit using the   MIN and   MAX functions:   =   MIN(G8 :   G19)
and
= MAX(G8   :   G19)
These functions provide   a quick and efficient way to   summarise your   data   and   gain insights   into   overall performance.
2.7          Plotting   a   Line   Graph   for   Monthly   Profit
Visualising data is a key part of analysis, as it allows you to easily identify   trends   and patterns.   Lastly,   we’ll   create   a   line   graph   to   display   the   monthly   profit   using   Excel’s   Insert   tab.
Step   1    Select   the   Data:   Highlight   the   range   of   cells   containing   the   profit   values,   e.g.,   G7:G19, including   the   headers.   Step   2    Open   the   Insert   Tab:   Go   to   the   Insert   tab, where   you’ll   find   various   charting   and   graphing   options.
Step 3    Choose a Line Graph:   Click the   Line   Chart icon   and   select the   basic   line   chart   option   from   the   available   choices.   Step   4    Customising   the   Chart:
After the graph is created, you can format it to make it more   informative   and   visually   appealing:
•    Add   Axis   Titles:   Navigate   to   Chart   Design   →   Add   Chart   Element   →   Axis   Titles, and   then   label   the   horizontal axis   as   “Month” and   the   vertical   axis   as   “Profit”   .
•    Title   the   Chart:   Provide   an   appropriate   title   for   the   chart,   e.g.,   Monthly   Profit   for   2020.
•    Customise   the   Line   and   Data   Markers:    To   customise   the   line’s    appearance, double-click   on   it   to   open   the   Format   Data   Series pane.    In   the   Fill      Line tab,   click   Marker,   expand    Marker   Options,   and   choose   Built-in to select the desired marker type.   You can also adjust the marker   size   for better   visibility.
•    Fade   Out   Gridlines   (Optional):    To    soften   the   gridlines,   double-click   on   any   gridline   to   select   them.       In   the   Format   Major   Gridlines   pane, under   the   Fill      Line   tab, adjust   the   transparency   level   to   your   preference.
•    Add a   Trendline:   A trendline can represent   the   overall   direction   of the   data   over   time.   To   add   a   trendline,   go   to   Chart   Design   →   Add   Chart   Element   →   Trendline,   and   select   Linear.
In this example, a trendline with a positive slope   is   added,   showing   that   monthly profits   are   trending   upward.
To   distinguish   the   trendline   from   the   data   line, you   can   add   a   legend:   Go   to   Chart   Design   →   Add   Chart   Element →   Legend, and   select   the   desired   location   (e.g.,   Bottom).This process offers a   simple way to visualise monthly profit   using   a   line   graph,   making   it   easier   to   spot   trends   and   patterns   over time.   Now that you’ve learned how to create and customise a line graph, try experimenting with other types of   charts, such   as bar or pie charts, to visualise different aspects of your data on   your   own.

         
加QQ：99515681  WX：codinghelp  Email: 99515681@qq.com
