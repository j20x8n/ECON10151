java c
ECON10151   Lecture 5 
Automating Tasks with   Excel   Macros 
Oct   2024
Learning Outcomes 
•    Be able to record and   run   macros   in   Excel to automate table formatting   and   create   charts.
•    Understand the principles of macros and apply   them   to   efficiently   handle   repetitive   tasks   in   Excel.
Introduction In   this   week’s   lecture, we   will   explore   macros   in   Excel, which   are   powerful   tools   designed   to   automate   repetitive   tasks.    A   macro   is   essentially   a   series   of   instructions   or   commands   that   can   execute   multiple   actions   with   a   single click.   Macros are created using Visual Basic for Applications (VBA), a programming language integrated   into   Excel that allows you to customise functionality.With   macros,   you   can   perform   a   variety   of   tasks,   such   as   formatting   cells,   performing   calculations,   and   generating charts, all while saving time and   reducing the   likelihood   of   errors.   To   get   started, the   first   step   is   to   enable   macros   in   Excel.   This will   allow   us   to   harness   their   full   potential   as we   delve   into   creating   and   using   them throughout this course.
1 Setting Up Macros 
The   Macro   feature   is   located   under   the   Developer   tab.      However,   the   Developer   tab   is   not   displayed   in   the   Excel   ribbon by default, but you can easily   add   it   to   access   macro-related   features.
• For   Windows   Users
1.    Click   on   the   File   tab.
2.    Select   Options.
3.    In   the   Excel   Options   window, choose   Customise   Ribbon.
4.    In   the   Customise   the   Ribbon   section, locate   the   Main   Tabs   list   and   check   the   box   next   to   Developer.
5.    Click   OK   to   apply   the   changes.
• For   Mac   Users
1.    Click   on   Excel   in   the   top   menu.
2.    Preferences.
3.    Choose   Ribbon    Toolbar.
4.   the Customise the   Ribbon section, find the   Main Tabs   list and   check the   box   next   to   Developer.
5.    Click   Save   to   confirm   the   changes.
Save the file: Remember   to   save   your   file   with   the    .xlsm   extension   to   ensure   that   macros   can   be   enabled   and   utilized.
2 Practice with Macro 
Dataset 
The   dataset   contains   monthly   sales   data   for   different   Lego   themes   from   October   to   December.      Each   table   provides   information on the   number of   units sold   (in thousands) and   revenue generated   (in   USD   millions).
• Theme:   The   specific   Lego   product   category   (e.g., Star   Wars,   Friends).
•    Units Sold: Number of units   sold during   the   month   (in   thousands).
• Revenue:   Total   sales   revenue   for   the   month   (in   USD   millions).Imagine   you   are   a   data   analyst working   for   the   Lego   Company.    You   have   been   assigned   the   following   tasks   to   enhance   the   sales   data   tables   for   the   last   quarter   report   of   2023.    This   means   you   need   to   complete   the   following tasks for October,   November,   and   December.
Main Tasks: 
1.    Format   the   Table:
• Insert   Title:
– Insert   a   new   row   at   the   top   of   each   table.
– Merge   the   cells   in   this   row   to   create   a   title   space.
– Enter the title:   ”Monthly Sales” and ensure   it   is centred within   the   merged   cells.
– Bold the title and apply an   appropriate colour   to   enhance   visibility.
• Format   Headers:
– Bold the   headers of each table to   ensure they   stand   out.
– Select a different shade of colour for the   header cells   to   further   distinguish   them   from   the   data.
• Format   Values:
– Insert   the   currency   dollar   sign   $   for   revenue   data.
2.    Highlight   the   schemes   with   the   top   10% of   units   sold   and   create   a   chart   displaying   the   units   sold   for   each Lego theme across the   months.
2.1 Task 1: Format Tables with Macro To complete these tasks, we first   need   to   apply   them   to   the   data   for   October.   Then,   we   can   repeat   the   same   actions for   November   and   December.    The   Macro will   help   us   record the   steps for   October,   and we   can   then   run the   Macro to automate the formatting   and chart creation for   November   and   December.
Step 1 Record Macro 
Go to the Developer tab and click on Record Macro.

Figure   1:   Record   Macro 
Then, set   up the   information for this   macro and   name   it   ’Format’   .   Click OK

Figure   2:   Name   Macro
From this point forward, everything you do in Excel will be recorded.
1.      Insert   Title:
-    Insert   a    New   Row:      Right-click   on   the   first   row   number   of   your   table   and   select Insert from   the   context   menu to add a   new   row at   the   top   of   the   table.
-    Merge   Cells:   Highlight the cells   in the   newly   inserted   row that   span   across   the   width   of   your   table.   On the Home tab, click on the Merge  Center button   in the   Alignment   group.
-    Enter   Title:    Click   on   the   merged   cell   and   type   "Monthly    Sales   .   "   ,   then   press   Enter.    Ensure   the   title   is centred   in the   cell.
-    Bold   the   Title:   Highlight   the   title   text.   Click   on   the Bold button   (or   press   Ctrl      +    B)   in   the   Font   group on   the Home tab.
-   Apply   Colour:   With   the   title   cell   still   selected,   click   the Fill Color button   (paint   bucket   icon)   in   the   Font group on the Home tab.   Choose an appropriate colour from   the   palette   to   enhance   visibility.
2.    Format   Headers:
-    Bold the   Headers:    Highlight the   header   row   of your table.    Click on the Bold button   (or   press   Ctrl   +    B)   in   the   Font   group   on   the Home tab.
-    Select a   Different Shade of Colour:   With   the   he代 写ECON10151 Lecture 5 Automating Tasks with Excel Macros
代做程序编程语言ader   row   still   selected,   click   the Fill Color button   in   the   Font group.   Choose   a different shade of   colour   from   the   palette   to   distinguish   the   header   cells   from the data.
3.    Format   Values:
-    Bold   the   Month   Numbers:    Selecting   data   (A3:A9)   in   the   ”Month”   column.    Next,   click   on   the   Bold button   in   the   Home   tab.
-    Insert   Currency   Dollar   Sign   for   Revenue   Data:
.    Instead   of   selecting   the   data   directly,   use   the   shortcut   for   this.    Click   on   cell   D3   and   then   press Ctrl + Shift + Down Arrow to   highlight   the   revenue   data   until   the   next   empty   cell.

.    Right-click   on   the   selected   cells   and   choose Format Cells from   the   context   menu.   In   the   Format Cells dialog box, select the Number tab.   Click on Currency or Accounting from the   list   on   the   left.   Choose the appropriate options for decimal places and currency symbol (ensure   it   is set to   $English   (United   States)).   Click OK to   apply   the   formatting.
Step 2 Run Macro 
Now we   need to   use this   macro to   help   us adjust the format   of the data   for   November.
1.    Go   to   the   worksheet Nov, then   click   on   the   Developer   tab.   Click   on   Macros   -   choose ”Format”,   and   then select Run.

Figure   3:   Run   Macro
Please find the   results below.   Are they correct?  

Figure   4:   Results
Note   that   the   headers   and   the   title   have   been   formatted.   In   the ’Month’ column, not   all   instances   of ’Novem-   ber’ are   bolded due to the   presence of   additional   rows   in   the   table.    However, the   revenue data   is   consistently formatted with the currency sign.    Despite the varying   number   of   rows,   the   same   formatting   has   been   applied   throughout.   This discrepancy arises from the different   methods used to select   the   data.
To   improve   this   formatting,   we   can   either   re-record   the   entire   sequence   to   update   the   Macro   or   edit   the   Macro by   modifying the code   behind   it.
Step 3 Edit Macro Macros   are   fundamentally   based   on   Visual   Basic   for   Applications   (VBA),   which   will   be   introduced   in   the   next lecture.   For now, consider that each action   or   step we   perform   in   Excel   corresponds to   a   line   of   code   that   can   be   interpreted   by the computer.   This   means we   have the option   to   edit   the   code   if   we   want   to   adjust   any   steps.
Let’s   take   a   closer   look   at   this   Macro.
• Go   to   the   Developer   tab   - Click   on   Macros   - choose ”Format” - Select Edit 
Now you are in the VBA editor, where you can   see   the   code   behind   Excel.   Please find   the   figure   below.  

Figure   5:   Code   for   Macro
Let’s focus on the code   related to formatting values.

•    Edit Code:   To ensure we employ   the   same   procedure   to   select   ’November,’ we can   change   the   code   for   selecting all cells that   include   November to:
Range("A3").Select 
Range(Selection, Selection.End(xlDown)).Select 
The   new code should be   shown   as   below,  

Figure   6:   Edited   Macro
•    Close   the   window, and   let’s   clear   the   formatting   applied   earlier.   Then, run   the   macro   again   to   see   whether it can apply the same formatting despite the varying   number   of   rows.
Now the format should be applied to the   other table,   regardless   of the   number   of   data   rows.
2.2 Task 2: Highlighting Information and Creating Charts with Macros Step 1: Record a New Macro Named ”Highlight” 
1.      Go   back   to   the   worksheet   ”Oct.”
2.    Go   to   the   Developer   tab, click   on ”Record   Macro,” and   name   it   ”Highlight.” Then   click   OK.
3.      Go   to   the   ”Units   Sold   (thousand)”   column,   click   on   cell   C3,   and   use      Ctrl+Shift+Down   Arrow      to   select the cells until   the   first   empty   cell.
4.    Goto   the   Home   tab, click   on ”Conditional   Formatting,”   select ”Top/Bottom   Rules,”   then   choose ”Top   10%”,   and click   OK.
5.    Stop   recording.
Step 2: Record a New Macro Named ”Create a chart” 
1.    Navigate   to   the   worksheet   named ”Oct.”
2.    Go   to   the   Developer   tab, click   on   Record   Macro, and   name   it ”Create a   chart.” Then, click   OK.
3.    Select   cell   B2,   then   hold   down      Ctrl+Shift+Down   Arrow      to   select   all   cells   in   columns   B   and   C   until   the first empty   cell.
4.      Go to the   Insert tab and select Clustered Column   Chart to   create   the   chart.
5.    Stop   recording
Step 3: Run Macros 
Now,   we   can   run   these   macros   or   assign   them   to   buttons   to   filter   or   visualize   data   for   the   other   months:   November and   December.
• Navigate   to   the   Developer   tab.
•    Click   on   Insert   to   add   a   button.    Assign   the   macro ”Highlight” to   the   button   and   rename   the   button   to ”Highlight” 
•    Click   on   Insert   to   add   a   button.   Assign   the   macro ”Create a chart” to   the   button   and   rename   the   button to ”Create a Chart.” 
• Copy   and   paste   these   buttons   into   the   worksheets   for ”November” and ”December” .
•    Run macros by   clicking   on these   buttons
Important Note: Check the Chart Data 
After creating the charts, check whether   they   display   the   correct   data   on   the   worksheets.    If the   charts   do   not   show the correct data, follow these instructions to   edit the   macro:
• Go   to   the Developer tab.
• Click   on Macros and   select ”Create a chart.” 
• Click   on Edit. 
•    Delete any specific sheet references   in the code to   ensure   it works   on   the   active   sheet.   Example:

•    Close the window and try running the   macro again.

         
加QQ：99515681  WX：codinghelp  Email: 99515681@qq.com
