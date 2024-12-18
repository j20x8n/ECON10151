java c
Lecture   2:   Excel   Solver 
ECON10151:   Computing   for   Social   Scientists 
September   29, 2024 Excel’s   Solver   is   a   versatile   tool   designed   to   help   users   find   optimal   solutions   to   complex   decision-making   problems.   Whether   you’re   allocating   resources   in   finance,   managing   supply   chains   in   logistics,   or   scheduling   operations,   Solver   allows   you to work within constraints and identify the best outcome,   such as   maximising profits   or   minimising   costs.For example, imagine you’re managing   a factory   and   need   to   determine   the   ideal   production   levels   of two   products,   given   a limited supply of materials and labour.   Solver can help you calculate the most efficient allocation that maximises profit   while   staying within your resource limits.In   essence,   Solver   takes   your   objective—like   increasing   profit   or reducing   expenses—and   tests   different   combinations   of   variables,   subject to the constraints you   set.   It   simplifies decision-making   where   trade-offs   are   involved,   ensuring   the   result   is   not   just mathematically sound but also practical for real-world scenarios.
1 How to Install Excel Solver 
Solver is an Excel add-in that doesn’t load   automatically   when   you   install   Excel,   but   it’s   easy   to   enable.   Whether   you’re   using   a Mac or Windows, the steps are straightforward, though   they   differ   slightly between   the   two   operating   systems.
1.1 Mac 
To install Solver on a   Mac,   follow   these   steps:
1.    Open   Excel.
2.    Click   on   the   Tools   menu   at   the   top.
3.    Select   Excel   Add-Ins.
4.    In the Add-Ins   available list, tick the box   for   Solver   Add-In.
5.    Click   OK.
6. Note: 
•    If Solver   Add-In isn’t listed, click   Browse to find   and install   it.
• If   you’re   prompted   to   install   the   Solver   Add-In, select   Yes.
7.    Once   installed, you’ll   see   the   Solver   button   under   the   Data   tab.
1.2 Windows 
To install Solver on Windows, follow these   steps:
1.    Open   Excel.
2.    Click   File   in   the   top-left   corner, then   select   Options.
3.    In   the   Excel   Options   window,   click   Add-Ins.
4.    At   the   bottom   of   the   window, next   to   Manage, ensure   Excel   Add-ins   is   selected, then   click   Go.
5.    Tick the box for Solver   Add-In in   the   Add-Ins   available list.
6.    Click   OK.
7. Note: 
• If   the   Solver   Add-In   is   missing   from   the   list, click   Browse   to   locate   and   install   it.
• If   asked   to   install   the   Solver   Add-In,   select   Yes.
8.    Once   installed, you’ll   find   the   Solver   button   in   the   Analysis   group   on   the   Data   tab.
2 A Worked Example: Diet Optimisation In this example,   we’ll explore   a practical   scenario   often   faced by   fitness   trainers   and   dieticians:   how   to   create   a   cost-effective   yet nutritionally balanced meal plan.   Using Excel Solver, we’ll navigate through the constraints of   this problem to optimise our   meal choices,aiming for the best nutritional outcome at the lowest cost.
Let’s   consider   four   meal   options:      Salad,   Protein   Shake,   Grilled   Chicken,   and   Pasta.    Each   of these   meals   offers   different   nutritional values and comes with a   specific   cost:

Salad 
Protein Shake 
Grilled Chicken 
Pasta 
Calories 
300 
250 
450 
600 
Protein (g) 
10 
30 
35 
12 
Fat (g) 
7 
3 
10 
15 
Cost ($) 
6.5 
5 
12 
8 
The challenge is to create a   meal plan   that meets specific nutritional goals while minimising the total cost.   In   this   case,   the goals   are:
• At   least   1800 calories   per   day,
• A   minimum   of   90 grams   of   protein,
• No   more   than   45 grams   of   fat.Here’s how   Solver comes into play.   We need   to   select   the   number   of   servings   of each   meal   (Salad,   Protein   Shake,   Grilled   Chicken, and Pasta) that together satisfy these nutritional requirements.   At the   same   time,   we   aim   to   minimise   the   total   cost   of   the meal plan.   This type of   problem is ideal for Solver, as it allows us   to   work   within   defined   constraints   (calories, protein,   and   fat limits) while optimising for cost.For example, you might have   a   client   with   a   fixed   daily   budget   but   also   the   need   to   maintain   certain   nutritional   standards.   By   inputting   the   nutritional   data   and   cost   for   each   meal   into   Excel,   Solver   can   identify   the   most   cost-effective   combination   that   achieves   the   target   nutritional   intake.    This   not   only   saves   time   but   also   ensures   that   the   plan   is   scientifically   backed   by   quantitative analysis.
3 How to Use Solver in a Nutshell To   effectively   use   Excel   Solver   for   optimisation   problems, follow   these   key   steps.   Solver   works   by   varying   the   values   of   specific variables   (behind   the   scenes)   within   the   limits   you   define   to   find   the best possible   solution   to   your problem.    Here’s   a   simple   guide to get   started:
1. Construct a Detailed Spreadsheet:    Start   by   organising   your   spreadsheet   with   all   relevant   data.       Make   sure   that   the   problem   components—like   costs,   nutritional   values,   or   other   important   factors—are   clearly   laid   out   so   that   Solver   can   interpret them correctly.
• Identify Decision Variables: Decision variables are   the   values   Solver   will   adjust   to   find   the   optimal   solution.   In   Excel,   these   are   also   known   as Changing Cells.    For   example,   in   the   diet   optimisation   problem,   the   decision   variables are the number of servings of each meal option.
• Define the Objective Function:   The objective function is   what   you   want   to   optimise,   such   as   minimising   cost   or   maximising profit.   In Solver, this is referred to   as the Set Objective.   In   our   diet   example,   the   objective   function   is the total cost of the meal plan, which we   aim   to   minimise.
• Incorporate Constraints: Constraints are   the   rules   or   limits   that   your   solution   must   follow,   such   as   nutritional   needs   or budget   limits.    These   ensure   that   Solver’s   solution   makes   sense   in real-world   situations.    In   our   case,   the   constraints   are the minimum and maximum nutritional goals,   like   needing   at   least   1800   calories   and   no   more   than 45 grams   of   fat.
2. Run Solver:    Once   your   spreadsheet   is   set   up   with   the   decision   variables,   objective   function,    and   constraints,   you’re   ready to run Solver.   Head to the   Data tab,   click   on   Solver,   and   it   will   begin   adjusting   the   decision   variables   within   your   constraints to find the best possible outcome.
3. Review the Solution:   After   Solver   has   finished,   it   will   display   the   optimal   solution   directly   in   your   spreadsheet.    This   will include the values for the decision variables that   best   meet   your   objective,   while   adhering   to   the   constraints.   At   this   point, you can check the results and ensure they are   sensible for your particular problem.By following these   steps, you can   confidently   use   Solver   for   various   optimisation   challenges,   whether   it’s   finding   the best   resource   allocation, balancing   a budget,   or creating cost-effective   diets.    Solver handles   the   complex   calculations,   leaving   you   to focus on analysing the results and making informed decisions.
4 Solving the Diet Optimisation Problem 
4.1 Setting up the Spreadsheet To   solve   the   diet   optimisation   problem   using   Excel   Solver, we   need 代 写ECON10151: Computing for Social Scientists Lecture 2: Excel SolverSQL
代做程序编程语言  to   organise   our   data   so   that   Solver   can   process   it   efficiently.   This   involves   defining   the   decision   variables,   setting   up   the   objective   function,   and   establishing   the   necessary   constraints.   Follow these steps to set   up   your   spreadsheet:
Step 1: Define the Decision Variables 
• In   cells   B2 :   E2, create   headings   for   each   type   of   food   (e.g.,   Salad, Protein   Shake,   Grilled   Chicken,   Pasta).
•    In   cells   B3 :   E3, enter   initial   trial   values   for   the   amount   of   each   food   to   include   in   the   meal   plan.   Make   sure   at   least   one of the values is greater than zero to allow Solver to work with   a   non-empty   starting point.
Step 2: Set up the Objective Function 
• Reference the   number   of   units   of   each   food   from   your   decision   variables   by   entering   =B3,   =C3,   etc.,   in   cells   B7 :   E7.
– It’s important to reference the number of   units rather than manually typing them.   By referencing, any   changes made   to   the   decision   variables   (in   B3   :   E3)   will   automatically   update   the   rest   of your   calculations.    This   not   only   saves   time but also reduces the risk of errors, ensuring consistency across your calculations.
• In   cells   B8 :   E8, input   the   cost   per   unit   for   each   food   item   (e.g.,   6.5 for   Salad,   5 for   Protein   Shake,   etc.).
•    To calculate the total cost of the meal   plan,   use   the   SUMPRODUCT function   in   cell   B10.   The   formula   will   look   like   this:   = SUMPRODUCT(B7 :   E7,   B8 :   E8)The SUMPRODUCT function multiplies   the number of   units of each food (in   B7 :   E7) by its respective cost (in   B8   :   E8),   and then   sums   the   results.    This   function   essentially   performs   element-wise   multiplication   of   B7 × B8,   C7 ×   C8,   and   so   on,   then adds them together.   The formula is equivalent to:
Total   Cost   = B7 × B8+C7 ×   C8+ D7 × D8+ E7 × E8
This gives the total cost of the diet based on the quantities you have   selected   for   each   food.
Step 3: Establish the Constraints 
•    Recreate   the   table   from   the   problem   statement,   listing   the   nutritional   information   (calories,   protein,   fat)   for   each   food   item in cells   B14   :   E16.
•    Use the SUMPRODUCT function again to calculate the total nutrients based on the amounts chosen in the decision variables.   For example, to calculate total calories, use:
= SUMPRODUCT($B$7 : $E$7,   B14 :   E14) 
This will give you the total calories consumed based on the   servings of each food. 
–    Note: The dollar   signs   ($$) in the formula   ensure   that   the   cell   references   remain   fixed   (absolute   references)   when   copying the formula to other cells.
• In   Column   G, specify   the   inequalities   for   your   constraints   (e.g.,   >=      1800 for   calories,   <= 45 for fat).
• In   Column   H, enter   the   target   values   for   each   constraint   (e.g.,   1800   for   calories,   45   for   fat).
By following these   steps, you will have   constructed   a   well-organised   spreadsheet   that   Solver   can   use   to   optimise   the   meal   plan.   Once everything is set up, you are ready to run   Solver   and find   the best   solution.
4.2 The Solver Parameters Dialog Box 
To effectively use Solver, follow these   steps:
1. Open the Solver Parameters Dialog Box 
Begin by navigating to   Data   >   Solver to open the Solver Parameters Dialog Box.
2. Set the Objective and Problem Type In   this   step:
•    In   the   "Set   Objective" box,   specify   the   cell   that   calculates   the   objective   function   (e.g.,   Cell   B10,   which   calculates the total cost).
•    Choose whether you want to minimise or maximise   the   objective.   For   our   problem,   since   we   want   to   minimise   the   total   cost,   select   "   Min".
3. Identify Decision Variables 
Next, specify the cells that represent your decision variables:
•    Click in the "By   Changing   Variable   Cells" box.
•    Select   the   cells   containing   the   decision   variables   (e.g.,   B3 :   E3).    These   are   the   cells   Solver   will   adjust   to   find   the optimal solution.
4. Add Constraints 
To ensure that Solver respects the constraints in the problem:
• Click   the   "Add" button   on   the   right.
•    In   the   "Cell   Reference" box,   select   the   cell   that   calculates   the   total   for   the   constraint   (e.g., for   calories,   choose   Cell F14, which   sums   the   total   calories   consumed).
•    Choose   the   appropriate   constraint   type   (<=, >=, =) and   then   input   the   target   value   for   that   constraint   (e.g., Cell   H14,   which   specifies   at   least   1800   calories).
•    Click   "OK" to   add   the   constraint.
Repeat this process for each constraint (e.g., protein, fat) to ensure   Solver respects   all nutritional requirements.
5. Make Variables Non-Negative 
Ensure all decision variables remain non-negative:
•    Check   the   box   titled    "Make    Unconstrained   Variables    Non-Negative".       This   ensures   that   all   variable   values   are   greater than or equal to zero, meaning Solver won’t   suggest negative   servings   of food.
6. Select the Solving Method Choose the solving method:
•    Select   "Simplex   LP" as the   solving method.   This method   is   appropriate   for   linear programming   problems   like   this one.
7. Solve the Problem 
Once everything is   set up:
•    Click the   "Solve" button to   run   Solver   and   find   the   optimal   solution.    Solver   will   adjust   the   decision   variables   and   provide a solution that minimises   the cost while meeting   all the   constraints.
4.3 Solution 
Once you’ve completed the Solver process, you’ll notice the following changes in   your   spreadsheet:
•    The Solver Parameters Dialog Box will close   automatically.
• The   values   in   the   decision   variable   cells   (e.g.,   B3 :   E3)   will   update   to   reflect   the   optimal   solution   that   Solver   has   found.
•    As   a result,   the   objective function   and   caluculations under   constraints in   your   spreadsheet   will   also   adjust   to   reflect   the   updated decision variables.
This updated information indicates that Solver has successfully applied the   optimal   solution to   your   optimisation problem.
4.4 Final Remark: 
In   practical   applications,   serving   sizes   are   usually   represented   as   whole   numbers   (e.g.,   you   can’t   eat   half a   serving   of grilled   chicken).   If the optimal solution contains fractional serving sizes, this might   appear   unusual.
Question: How might we adjust our constraints or model to ensure   that the   resulting   serving   sizes   are   whole   numbers?
Answer: To ensure that Solver produces whole numbers for the   serving   sizes,   you   need   to   adjust   the   constraints   to   require   integer values for the decision variables.   Follow these steps:
• Open   the   Solver   Parameters   Dialog   Box   again   by   clicking   on   Data   >   Solver.
• Click   the   "Add" button   to   introduce   a   new   constraint.
• In   the   "Cell   Reference" box,   select   the   cells   representing   the   decision   variables   (e.g.,   B3 :   E3).
•    In   the   "Constraint"   box,   select   int   (integer),   which   ensures   that   the   values   for   the   decision   variables   are   restricted   to   whole numbers.
• Click   "OK" to   add   the   constraint, and   then   click   "Solve"   again   to   re-run   Solver   with   the   updated   settings.
By adding this integer constraint, you ensure   that   Solver provides   solutions   with   whole   number   servings,   which   makes   the   results more practical   for real-life applications.

         
加QQ：99515681  WX：codinghelp  Email: 99515681@qq.com
