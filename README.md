java c
Lecture   7:   Advanced   Data   Analysis   in   Excel   using   ToolPak
ECON10151:   Computing   for   Social   Scientists
November   10, 2024In previous lectures, we covered essential   Excel   tools   for   organising   and   summarising   data.   We   manually   calculated   mea-   sures   like   the   mean, variance, and   standard   deviation, and   used   Pivot   Tables   for   flexible   data   summaries.   Today, we’ll   take   this further   by   introducing   a   more   efficient   method   — the   Excel   Analysis   ToolPak.The   ToolPak is   an   add-in   that   performs   statistical   calculations   quickly   and   accurately,   allowing   you   to   run   analyses   with   just   a few   clicks.   This is   especially valuable for   larger   datasets,   saving   time   and   reducing   errors.    The   ToolPak provides   tools   for advanced data analysis, including statistical tests, visualisations,   and regression,   all within   Excel.
We’ll   start   by   using   the   ToolPak   to   analyse   student   performance   data   from   a   sample   dataset   with   scores   of   72   students   across different subjects.   The dataset includes:
•    Student   ID
•      Gender
•    Math   score
•    Writing   score
1          Installation   Guide:   Setting   Up   the   Data   Analysis   ToolPak
To get started, let’s make sure the   ToolPak is enabled   in   Excel.   Follow   the   steps   below   for   your   specific   operating   system.
1.1          Mac
To enable the Analysis   ToolPak on   a   Mac:
1.    Open   the   Tools   menu.
2.    Select   Excel   Add-Ins.
3.    Tick   the   Analysis   ToolPak   checkbox   and   click   OK.
4.    If   the   Analysis   ToolPak   is   not   listed, click   Browse   to   find   it,   or   select   Yes   if   prompted   to   install   it.
5.    Once   installed, the   Data   Analysis   button   will   appear   on   the   Data   tab.
1.2          Windows
To   enable   the   Analysis   ToolPak   on   Windows:
1.    Navigate   to   File   >   Options   >   Add-Ins.
2.    In the   Manage dropdown box, choose   Excel   Add-ins and click   Go.
3.    Tick   the   Analysis   ToolPak   checkbox   and   click   OK.
4.    If   the   Analysis   ToolPak   is   not   available, click   Browse   to   locate   it,   or   select   Yes   if   prompted   to   install   it.
5.    Once   installed, the   Data   Analysis   option   will   appear   in   the   Analysis   group   on   the   Data   tab.
1.3          ToolPak   Overview
Clicking on   Data   Analysis opens a dialog box   with   a   variety   of tools for   performing   data   analysis   using   built-in   mathematical   formulas.
Here are the key tools we will be using   today:
•    Descriptive Statistics
•    Rank and   Percentile
•      Correlation
•    Regression
These tools   are   crucial for identifying   data patterns,   summarising   information,   and   supporting   informed   decision-making.   If some of these terms   are new, don’t worry — we will cover each   tool   in   detail   with practical   examples.
2          Descriptive   StatisticsThe   Descriptive   Statistics   tool   is   one   of the   simplest   yet   most   powerful   options   for   summarising   data.    This   tool   provides   a   quick   overview   by   producing   essential   summary   measures,   such   as   the   mean,   median,   and   variance.      Let’s   use   this   tool   to   analyse the math score variable in our   dataset.
Once the Analysis   ToolPak is enabled, follow these steps to   generate   descriptive   statistics in   Excel:   Step   1:    Open   the   Data   Analysis   Dialog   Box:
•    Mac:   Go to   Data   >   Data   Analysis.
•    Windows:   On   the   Data   tab, in   the   Analysis   group, click   Data   Analysis.   Step   2:    Select   Descriptive   Statistics   from   the   list   and   click   OK.
Step   3:    In   the   Descriptive   Statistics   dialog   box, set   the   following   options:
(a)    Input   Range:   Select   the   range   of   data   to   analyse.   For   our   example, choose   the   math   scores   in   C1:C73.
(b)    Grouped   By:   Select   whether   your   data   is   organised   by   columns   (default) or   rows. For   this   dataset, keep   "Columns"   selected.
•    Choose   "Rows" only   if   your   data   is   arranged   horizontally.
(c)    Labels   in   First   Row:   Tick   this   box   if   your   data   includes   column   headers   in   the   first   row.
(d)    Output   Range:   Specify   where   to   display   the   results.   You   can:
•    Place the output in a new worksheet to keep things   organised,   or
•    Select a specific cell, such   as   H1,   in   the   current   worksheet,   ensuring   there   is   enough   space   for   the   output.   Note:   The output requires at least 2 columns per variable, so   make   sure   there   is   adequate   space.
(e)    Tick   Summary   Statistics   to   generate   key   measures.   (f)    Click   OK   to   produce   the   table.
The report will include important statistics   such as:
•    Central Tendency:   Measures that summarise the centre of a dataset, including the   mean   (average   value),   median   (mid-   dle value in sorted data),and mode (most   common value).
•    Variability:   Statistics that show how spread out the   data is,   including:
–    Standard Deviation and Variance:   Indicate the dispersion of data points around the mean.
–    Minimum and Maximum:   The smallest and largest values   in the   dataset.
–    Range:   The difference between the maximum and minimum, indicating the total   spread.
•    Sum and Count:   The total of all values and the   number   of data points.Note:   Kurtosis and Skewness are also shown to   indicate   the   shape   of   the   data   distribution;   the   Standard Error   shows   how   much the average value (mean) might vary if we took different samples. It helps indicate how   precise the mean is.   These details   are not crucial to know   for this course.
2.1          Analysis   Limitations:   Text   Data
Attempting   to   generate   a   descriptive   statistics   table   for   the   Gender   variable   will   result   in   an   error   message:    “Descriptive   Statistics   -   Input   range   contains   non-numeric   data.”
Text data can be difficult to analyse quantitatively, so   it   often   needs   to   be   recoded   into   a   numerical   format:
•    Binary   Text   Data   (e.g., yes/no, true/false) can   be   converted   to   0s   and   1s   for   easier   analysis.
•    Ordered Categories can sometimes be mapped to integers.   Examples   include:
–    Freshman, Sophomore, Junior, Senior
–    Strongly Agree, Agree, Disagree, Strongly Disagree
•    Some   text   data   may   not   translate   meaningfully   into   numbers.    For   instance,   country   names   cannot   be   easily   ranked   or   quantified.
In   our dataset, the   Gender column is   binary   (female/male),   so   we   can   recode   it   as   1s   and   0s   using   the   IF function.    Recall   that this function performs a conditional test, returning one value if the condition is TRUE   and   another   if it is FALSE:
= IF(logical_test,value      if      true,value      if      false)
To   create   a   binary   numerical   variable   for   Gender, enter   the   following   formula   in   a   new   column   (e.g.,   cell   E2):   = IF(B2 = “female”   ,   1,   0)
This assigns a value of   1 if the student is female   (as indicated   in   B2)   and   0   otherwise.   Label   Column   E   as   Gender   Dummy.Once recoded, we can generate descriptive statistics for   all three   variables   in   the   dataset   by   setting   the   Input   Range in   Step
3 above   to   C1:E73.
3          Rank   and   PercentileThe   Rank   and   Percentile   tool   in   the   Analysis   ToolPak   helps   quickly   identify   the   rank   of   values   in   a   list   and   the   corresponding percentile   for   each   value.      The   percentile   indicates   the   percentage   of data   points   that   fall below   a   given   number,   showing   the   relative position of   each data point within the dataset.
To   illustrate, let’s   calculate   the   Rank   and   Percentile   for   the   writing   score   data:   Step   1:    Open   the   Data   Analysis   Dialog   Box:
•    Mac:   Go to   Data   >   Data   Analysis.
•    Windows:   On   the   Data   tab, in   the   Analysis   group, click   Data   Analysis.
Step 2:    Select   Rank   and   Percentile from the list   and   click   OK.
Step   3:    In   the   Rank   and   Percentile   dialog   box,   set   the   following   options:
(a)    Input   Range:   Select   the   range   for   the   writing   scores   (e.g.,   D1:D73).   (b)    Grouped   By:   Ensure it   is   set   to   Columns.
(c)    Labels   in   First   Row:   Tick   this   box   since   the   first   row   contains   column   headers.
(d)    Output   Range:   Choose   where   to   display   the   results   (e.g.,   O1),   ensuring   there   is   enough   space   for   the   output.   (e)    Click   OK   to   generate   the   table.
The output table includes four columns:
•    Point:   The position of each value in the original list, allowing you to match   values   to   their   original   order.
•    Writing   Score:   The   original   data   values   (e.g., writing   scores), retaining   the   original   label.
•    Rank:   The rank of   each writing score, sorted in descending   order,   shows how   each   score   compares   within   the   dataset.
For   example, the   highest   score   will   have   a   rank   of   1, the   next   highest   will   be   2, and   so   on.   This   helps   you   quickly   identify where each value stands in relation to   others.
Note:   Scores with the same value will share the   same   rank.
•    Percent:   The percentile rank indicates the percentage of data points that   fall   below   each   writing   score.   This   helps   show   the relative standing of each score within the   dataset.
For instance:
–    If   a   writing   score   is   in   the   100th   percentile,   100% of   the   scores   in   the   dataset   are   lower   than   this   value   —   this   score will have the highest rank.–   代 写ECON10151: Computing for Social Scientists Lecture 7: Advanced Data Analysis in Excel using ToolPakProcessing
代做程序编程语言 If   a   writing   score   is   in   the   50th   percentile, 50%   of   the   scores   in   the   dataset   are   lower   —   this   represents   the   median.
4            Correlation
Loosely   speaking,   correlation   measures   how   strongly   two   variables   are   related,   indicating   whether   they   move   together   in   a   similar way:
•    A positive correlation means that   as one   variable   increases,   the   other   tends   to   increase   as   well   (or   if   one   decreases,   the   other also decreases).   For example, as study time goes   up,   test   scores   might   also   go   up.
•    A negative correlation means that as one   variable   increases,   the   other tends   to   decrease.   For   instance,   as   the   number   of   hours spent watching Netflix increases, time spent studying   might   decrease.
•    Correlation   takes   values   between   -1   and   1:
–    A   correlation   value   close   to   ±1   indicates   a   strong   linear   relationship   between   the   variables,   meaning   they   move   closely in sync, either in the same direction   (positive correlation)   or   in   opposite   directions   (negative   correlation).
–    Values near 0 suggest little to no linear relationship between the   variables.
(You will learn about correlation more   formally in the Semester   2   Advanced Statistics course.)
To explore the relationship between variables, follow these steps:   Step   1:    Open   the   Data   Analysis   Dialog   Box:


•    Mac:   Go to   Data   >   Data   Analysis.
•    Windows:   On   the   Data   tab, in   the   Analysis   group, click   Data   Analysis.   Step   2:    Select   Correlation   from   the   list   and   click   OK.
Step   3:    In   the   Correlation   dialog   box, set   the   following   options:
(a)    Input   Range:   Select   the   range   for   the   data   you   want   to   analyse   (e.g.,   C1:E73 for   the   scores   and   gender   dummy).   (b)    Grouped   By:   Ensure it   is   set   to   Columns.
(c)    Labels   in   First   Row:   Tick   this   box   if   the   first   row   contains   column   headers.
(d)    Output   Range:   Choose   where   the   results   should   be   displayed   (e.g.,   H22), ensuring   there   is   enough   space   for   the output.
(e)    Click OK. Excel will generate a table   showing the   correlation   coefficients between   the   variables.
5          RegressionRegression analysis helps us identify trends and understand relationships between variables.   Today, we’ll use Excel to perform   simple regression   analysis   on   a   dataset from   Starbucks.    The   dataset includes   annual   advertising   costs from   2000 to   2018   in   column   B   (input   variable, X) and   sales   revenues   in   column   C   (output   variable, Y).   Although   many   factors   affect   sales,   we   will focus on these two variables for   simplicity.
Previously, we used scatter plots to visualise the relationship between variables by plotting data points.
Today,   we’ll   take   this   further   by   applying   simple   linear   regression,   which   fits   a   line   to   the   data   to   quantify   the   rela-   tionship   and   make   predictions.    Specifically,   we’ll   model   the   relationship   between   advertising   costs   (X)   and   sales   revenue   (Y).
Our   objective   is   to   learn how   to use   simple   linear regression   to predict   sales   revenue   from   advertising   costs.    In   short,   we   want to answer:   If we know the advertising cost, can we predict sales revenue,   and   how?
The linear regression   model is expressed as:
Y =   a+bX   +   e,
where:
•   X:   The   input   variable   (advertising   cost).
•    Y:   The   output   variable   (sales   revenue).
•    a:   The   Y-intercept, or   the   estimated   Y   value   when   X   is   0.
•    b:   The slope, which tells us how much Y   is predicted to change   for   a   one-unit   change   in X.
•    a+bX:   The equation used to predict Y based on X.•    e: The   error   term, or   prediction   error   — the   difference   between   actual   and   predicted   Y   values, also   known   as   the   residual.
The e term indicates that our predictions may not be perfect due   to   factors   not   included   in   the   model.   Usually,   e   is   not   zero   because other variables, such as customer preferences or   store locations,   can   influence   sales revenue.
5.1          Creating   a   Scatter   Plot   with   a   Trendline   in   Excel
Let’s revisit how to create a scatter plot and   add   a trendline to   visualise   a   simple   linear regression.   Follow   these   steps:Step   1:    Select   the   Data:    Highlight   the   data   range   for   the   input   (X)   and   output   (Y)   variables.    For   example,   select   the   range
B1:C20 to   include   both   advertising   costs   and   sales   revenues.   
Step   2:    Insert   the   Scatter   Plot:




•    Go   to   the   Insert   tab   at   the   top   of   the   Excel   window.
•    In   the   Charts   group, click   on   the   Scatter   icon   and   select   Scatter   with   only   Markers.
Step   3:    Add   Labels   and   Titles:
•    Click   on   the   chart   to   activate   the   Chart   Design   Tools.
•    Click   the   Add   Chart   Element   icon   (the   + sign)   and   add   Axis   Titles   and   a   Chart   Title.
•    Label the x-axis as “Advertising   Costs   (X)”   and the y-axis   as “Sales   Revenue   (Y)”   .
•    Edit the chart title to a   descriptive   name,   such   as “Relationship   Between   Advertising   Costs   and   Sales   Revenue”   .
Step   4:    Add   a   Trendline:
•    Right-click   on   any   data   point   in   the   scatter   plot   and   select   Add   Trendline.
Alternatively, you   can   click   the   Add   Chart      Element   icon   (the   + sign)   and   choose   Trendline   from   the   dropdown menu.
•    Choose   Linear   as   the   trendline   type.
•    Tick   the   Display   Equation   on   chart   box   to   show   the   regression   equation.
•    To make the trendline more visible, change the line colour to red   and   adjust   the   line   style   if desired.Explanation: The trendline   represents the   best-fit straight line through the data   points (in the sense of   minimising   prediction   errors),   illustrating   the   linear   relationship   between   advertising   costs   and   sales   revenue.         The   displayed   equation   (e.g.,   y   =   1.1343x −   8.0544) is   your   regression   line, where   a   = −8.0544 is   the   intercept   and   b   =   1.   1343 is   the   slope.
5.2          The   Regression   Tool   in   ToolPak
We   can   take   this   regression   analysis   further   with   the   Regression   Tool   in   the   Analysis   ToolPak.   This   tool   allows   us   to   gain   more detailed insights into the relationship between variables by providing key statistical outputs.
To perform. regression analysis in Excel:         
Step   1:    Open   the   Data   Analysis   Dialog   Box:
•    Mac:   Go to   Data   >   Data   Analysis.
•    Windows:   On   the   Data   tab, in   the   Analysis   group, click   Data   Analysis.   
Step 2:    Select   Regression from the   list   and   click   OK.
Step   3:    In   the   Regression   dialog   box, set   the   following   options:
(a)      Input   Y   Range:   Select the output variable, e.g.,   Sales    Revenues (C1:C20).
(b)    Input   X      Range:    Select   the   input   variable,   e.g.,   Advertising   Costs   (   B1:B20).    (If   there   are   multiple   X   variables,   they should be in adjacent   columns.)
(c)    Tick   the   Labels   box   if   your   data   has   headers.
(d)    Choose   where   to   display   the   output,    either   in   a   new   worksheet   or   in   a   specific   cell   (e.g.,    E1)   in   the   current worksheet, ensuring there is enough space for the results.
(e)    Optional   settings:
•    Check   Residuals   to   see   the   differences   between   predicted   and   actual   values.
•    Check   Line   Fit   Plots   to   visualise   actual   versus   predicted   values.
•    Check   Residual   Plots   to   visualise   the   residuals.   
(f)    Click   OK   to   generate   the   regression   analysis   output.


Key Insights from Regression Outputs:
•    The coefficients show the values of a (intercept) and   b   (slope),   which   define   the regression   line   equation.
•    The   residual   output   includes   the   predicted   Y   values, calculated   as   a+bX   for   each   X   , and   the   residuals, e =   Y −(a+bX),   which are the differences between actual and predicted Y   values.
•    The line fit plot shows the actual data   alongside   the predicted   Y   values,   similar   to   a   scatter plot   with   a   trendline.To   customise   the   marker   format,   click   any   marker   in   the   plot   to   open   the   Format      Data   Series   pane.      Then,   click   the   Marker button, expand the   Marker   Options dropdown,   and change the   setting   from   Automatic to   Built-in to   modify   the   marker   style.
•    The residual plot visualises the residuals,   indicating   where   the   predicted   line   deviates   from   actual   data   points.   Positive   and negative residuals show over- and under-predictions, respectively.Note:    The   output   may   include   standardised   residuals   and   a   normal probability plot,   which   standardise   the   residuals   by   their mean and standard deviation to check if they are normally distributed.   For this course, you do not need to know these two   in   detail.
Take-Home Exercise:
Use   the   Regression   Tool to   find   the   line   that best predicts   the   Math   score based   on   the   Writing   score   and   Gender   from   the student performance dataset.Hint:   In   Step   3   above,   for   the   Input   Y   Range,   select   the   Math   score   (C1:C73),   and   for   the   Input   X   Range,   select   both   the   Writing   score   and   Gender   Dummy   (D1:E73).    (You   do   not   need   to   generate   the   line   fit   plot   or   the   residual   plot   for   this exercise.)

         
加QQ：99515681  WX：codinghelp  Email: 99515681@qq.com
