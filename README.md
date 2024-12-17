java c
ECON10151   Lecture 4
Managing   Data with   Excel   Functions
Oct   2024
Learning Outcomes
•    Be   able to   effectively   retrieve data from   a dataset   using   VLOOKUP,   INDEX,   and   MATCH   functions,   and   apply them to   large datasets.
•    Understand the   differences   and   use   cases   of   VLOOKUP versus   INDEX   and   MATCH,   enabling them   to   choose the appropriate function based on specific data   retrieval   needs.
IntroductionExcel   offers   a   wide   range   of   powerful   functions   that   can   help   you   manage,   analyse,   and   extract   valuable   insights from data.   This week, we’ll focus on three key functions that   are   essential for working with   data   effec-   tively:   VLOOKUP,   MATCH,   and   INDEX.   These   functions   will   enable   you   to   quickly   find   specific   information,   locatedata within a table, and   retrieve values from different parts of   your   dataset.
1         VLOOKUP
In   this   section,   we   will   explore   the   syntax   of   VLOOKUP   and   will   break   down   the   arguments   so   that   you   can understand   how they work.
1.1         VLOOKUP   Syntax
The   VLOOKUP   function   searches   for   a   value   in   the   first   column   of   a   table   and   returns   a   value   in   the   same   row from another column.
Syntax:
=VLOOKUP(lookup   value,      table   array,    col         index   num,      [range   lookup])
Arguments:
•    lookup   value:   The value you want to search for in the first column   of the table.
•   table   array:    The   range   of   cells   that   contains   the   data.      The   first   column   of   this   range   is   where   the   lookup   value will be searched.
•    col         index   num:   The   column   number   (starting   from   1) from   which   you   want   to   retrieve   the   value.
•      [range   lookup]   : An   optional   argument.   Use   TRUE   for   an   approximate   match   or   FALSE   for   an   exact   match.   FALSE   is   recommended   for   most   cases   to   ensure   accurate   results.
1.2            Practice   with   VLOOKUP
InstructionsIn   this   exercise, you   will   practice   using   the   VLOOKUP   function   in   Excel   to   retrieve   specific   data   from   a   table.   You will be working with a dataset that contains information about students, their departments, academic year, and   GPA.   Follow the tasks below to complete the   exercise.
Dataset
The dataset contains the following columns:
•      Student   ID:   Unique   identifier for each student.
•    Name:   Name   of   the   student.
•    Department:   The department the student belongs to.
•   Year:   The academic year of the   student.
•    GPA: The   student’s   Grade   Point   Average.
Task   1:   Find   GPA   for   Given   Students
You are provided with the   names   of two   students:
•   Alice   Smith
•    George   Patel
Your   task   is   to   use   the   VLOOKUP   function   to   find   their   GPAs   from   the   dataset.   Follow   these   steps:
1.    Open   Excel L4   Data and work on the worksheet   named   as   VLOOKUP.
2.    In cells   B15 and   B16,   use the VLOOKUP   function   to   find   the   GPA   for   each   of   the   two   students.
3.    Syntax   of   VLOOKUP:
=VLOOKUP(lookup   value,      table   array,    col         index   num,      [range   lookup])
4.      Setup
•    lookup   value:   The value you want to search for in the first column of the table.
- Set   it   to   the   student’s   name.
•   table   array:   The range of cells that contains the data.   The first column   of this   range   is where   the   lookup   value will be searched.
- Select the data   range covering the columns from   Student   Name to   GPA.
•    col         index   num:   The column   number   (starting from   1) from which you want to   retrieve the value.
-   It   is   the   GPA   column, which   is   the   4th   column   in   the   range.
•      range   lookup:   An optional argument.   Use TRUE for an approximate   match   or   FALSE for   an   exact   match.
- We set   it to FALSE to   ensure   an   exact   match   is   found.
Remember:      Make   sure   the   lookup   value   is   in   the   first   column   of   the   table   array.      Therefore,   when   you   select data for the table array, ensure that the   student   names   are   in the first   column   of   this   array.
The   GPA   for   Alice   Smith:             = VLOOKUP(A15,      $B$2:$E$11,    4,      FALSE)
The   GPA   for   George   Patel:             = VLOOKUP(A16,    $B$2:$E$11,      4,      FALSE)
5.      Graphical   illustration of the formula:
   
Figure   1:   Task   1 VLOOKUP   Example
6.    Report   the   GPAs   for   Alice   Smith   and   George   Patel.    You   should   obtain   the   following   results   once   you   have entered the formula.
   
Figure   2:   Task   1   Results
Task 2:   Check   if Student   Name   Exists   in the   Dataset
You are given the   following   names:
•   Julia   Fernandez
•    Michael   Green
Your   task   is   to   use   the   VLOOKUP   function   to   check   if   these   names   exist   in   the   dataset.   Follow   these   steps:
1.   Think:   How   to   use   VLOOKUP
If   the   name   exists,   return   the   corresponding   Name   from   the   list.    If   not, the   function   should   return   an   error (such   as   #N/A).
2.    Syntax   of   VLOOKUP:
=VLOOKUP(lookup   value,      table   array,    col         index   num,    FALSE)
3.      Setup
•    lookup   value:   The value you want to search for in the first column of the table.
- Set   it   to   the   student’s   name.
•   table   array:   The range of cells that contains the data.   The first column   of this   range   is where   the   lookup   value will be searched.
-   Define table   array as   the data range covering the single column   of   Student   Name.
•    col         index   num:   The column   number   (starting from   1) from which you want to   retrieve the value.
- Assign   col         index   num to the first column, which is the only   column   in the   range.
•      range   lookup:   An optional argument.   Use TRUE for an approximate   match   or   FALSE for   an   exact   match.
- Set   range   lookup   to   FALSE   to   ensure   that   an   exact   match   is   found.
Remember:      Make   sure   the   lookup   value   is   in   the   first   column   of   the   table   array.      Therefore,   when   you   select data for the table array, ensure that the   student   names   are   in the first   column   of   this   array.
Julia   Fernandez:      代 写ECON10151 Lecture 4 Managing Data with Excel FunctionsR
代做程序编程语言       = VLOOKUP(A20,$B$2:$B$11,1,FALSE)
Michael   Green:             = VLOOKUP(A21,$B$2:$B$11,1,FALSE)
4.      Graphical   illustration of the formula:
   
Figure   3:   Task   2 VLOOKUP   Example
5.    Report whether   each   of   the   two   names   is   found   in   the   dataset.    You   should   obtain   the   following   results   once you   have entered the formula.
   
Figure   4:   Task   2   Results
This suggests that Julia   Fernandez   is   included   in the dataset;   however,   Michael Green   is   not   on   the   list,   as   it   returns   an   error   (#N/A   )in   the   cell.
2         INDEX and   MATCHIn this section, we will   briefly explore the   syntax   of the   INDEX   and   MATCH   functions.   The   INDEX   function   returns   a   value   based   on   specified   row   and   column   numbers   within   a   given   array,   while   the   MATCH   function   finds   the position of a value   in   a   row   or   column.
We will break down their arguments to   help you   understand   how each   function works.
INDEXThe   INDEX   function   returns   the   value   in   a   specified   cell.    You   need   to   know   the   row   and   column   numbers   of   the   cell   within   an   array   so   that   Excel   can   locate   its   position.    Once   identified,   the   function   returns   the   value   contained in   that   cell.
Syntax:
=INDEX(array,      row   num,      [column   num])
Arguments:
•    array:   The   range of cells that contains the data from which you want to   retrieve   a value.
•      row   num:   The   row   number   in   the   array   from   which   you   want   to   retrieve   a   value.
•      [column   num]   :   The   optional   column   number   in   the   array.   If   omitted, the   function   will   return   the   value   from the first column.
MATCH
The   MATCH   function   searches   for   a   specified   value   in   a   specific   column   or   a   row   and   returns   the   relative   position   of   that   value.
Syntax:
=MATCH(lookup   value,      lookup   array,      [match   type])
Arguments:
•    lookup   value:   The value you want to search for in the   array.
•    lookup   array:   The   range   of   cells   that   contains   the   data   you   want   to   search.    The   lookup   array   here   can only be a   single   column   or   a   single   row.
•      [match   type]   :   An optional argument.    It specifies   how   Excel should   match the   lookup   value.    Use   1   for   the   largest   valueless   than   or   equal   to   the   lookup   value, 0 for   an   exact   match, or   -1 for   the   smallest   value greater than or equal to the   lookup value.   For   most scenarios, 0   is   preferred   to   ensure   an   exact   match.


2.1            Practice   with   INDEX   and   MATCH
Task   1:   Obtain the GPA of the Student with   ID   33104   Using   INDEX   and   MATCHManually   finding   a   student’s   GPA   would   involve   identifying   the   correct   row   where   the   Student   ID   appears   and then locating the GPA in the corresponding column.   For small datasets, this is   simple,   but   as datasets   grow   larger, manually   finding   this   information   becomes   prone   to   error   and   time-consuming.   With   Excel’s   INDEX and   MATCH   functions, we   can   automate   this   task   and   ensure   accuracy.   Thus,   in   this   task,   we   can   use   MATCH to find   out the   row   number   and the   column   number.    Then,   INDEX   can   return the   value   of   GPA   based   on the   row   number of Student   ID and the   column   number   of   GPA.
Your   objective   is   to   find   the   GPA   of   the   student   with   the   Student   ID:   33104.    To   achieve   this   goal,   we   can follow these steps:
1.    Use   the   MATCH   function   to   find   the   row   number   where   the   Student   ID   33104   is   located.
•    (Worksheet:   MATCH   and   INDEX)   Begin   with   the   cell   I7 and   type   MATCH   formula   in   the   cell.
•    Syntax   of   MATCH:
=MATCH(lookup   value,      lookup   array,      [match   type]).
•      Setup
。lookup   value:   The value you want to search for in   the array.
- This   case,   it   is   33104   (the   Student   ID   you’re   looking   for).
。lookup   array:   The range of cells that contains the data you want to search.   The lookup   array   here can only be a   single   column   or   a   single   row.
- We select the column of Student   IDs   (column   A   in   the   dataset).
。   [match   type] should   be   set   to   0 for   an   exact   match.
The   row   number   of   Student   ID   33104:
= MATCH(H7,    A2:A11,    0)
•      Graphical   illustration of the   formula:
   
   
Figure 5:   Task   1   MATCH   Example:   row   no.
   
2.    Use   the   MATCH   function   to   find   the   column   number   where   the   GPA   is   located.
•    Begin   with   the   cell   I11 and   type   MATCH   formula   in   the   cell.
•    Syntax   of   MATCH:
   
=MATCH(lookup   value,      lookup   array,      [match   type]).
•      Setup



。lookup   value:   The value you want to search for in   the array.
-   It   is   GPA   (the   header   you   are   looking   for   in   the   row,   including   all   headers).
。lookup   array:   The range of cells that contains the data you want to search.   The lookup   array   here can only be a   single   column   or   a   single   row.
- We   select   the   row   of   headers   (Row   1   in   the   dataset).   。   [match   type] should   be   set   to   0 for   an   exact   match.
The column   number of   GPA:
= MATCH(H11,A1:E1,0)
•      Graphical   illustration of the   formula
   
Figure   6:   Task   1   MATCH   Example:   column   no.
3.    Use   the   INDEX   function   to   return   the   GPA   of   the   student   with   ID   33104
We can fill in the INDEX function based on the row number   and the column   number found   using the MATCH   function.
•    Begin   with   the   cell   I17 and   type   INDEX   formula   in   the   cell.
•    Syntax   of   INDEX: =INDEX(array,    row   num,       [column   num])
•      Setup
。array contains all of the values in this   student   information data   set.   。row   num   corresponds   to   the   row   number   of   the   student   ID   33104.
。col   num corresponds to the column   number where GPA   is   located.
The   GPA   of   the   student   with   ID   33104:
= INDEX(A1:E11,I7,I11)
•      Graphical   illustration of the   formula:
   
Figure   7:   Task1   INDEX   Example

         
加QQ：99515681  WX：codinghelp  Email: 99515681@qq.com
