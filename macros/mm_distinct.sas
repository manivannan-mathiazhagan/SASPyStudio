***-------------------------------------------------------------------------------------------------***;
*** Macro Name:    mm_distinct.sas                                                                  ***;
*** Purpose:       Display distinct values of one or more variables with occurrence counts.         ***;
***-------------------------------------------------------------------------------------------------***;
*** Programmed By: Manivannan Mathialagan                                                           ***;
*** Created On:    16Sep2026                                                                        ***;
***                                                                                                 ***;
***-------------------------------------------------------------------------------------------------***;
*** Note:          This macro is intended for personal programming, development, and debugging      ***;
***                use only and is not a study-standard or production macro.                        ***;
***-------------------------------------------------------------------------------------------------***;
*** Parameters                                                                                      ***;
***-------------------------------------------------------------------------------------------------***;
*** Name     | Description                                            | Default value   | Required  ***;
***----------|--------------------------------------------------------|-----------------|-----------***;
*** DS       | Input SAS dataset                                      | No default      | Yes       ***;
*** VARS     | Variable list defining the distinct combinations       | No default      | Yes       ***;
*** WHERE    | WHERE condition used to filter observations            | Blank           | No        ***;
*** OUT      | Output dataset of values and counts                    | _MM_DISTINCT    | No        ***;
***-------------------------------------------------------------------------------------------------***;
*** Output(s):      Dataset specified in OUT= and PROC PRINT output.                                ***;
*** Dependencies:  Variables specified in VARS= must exist in DS=.                                  ***;
***-------------------------------------------------------------------------------------------------***;
*** Modification History                                                                            ***;
***-------------------------------------------------------------------------------------------------***;
*** Date        | Modified By            | Description                                              ***;
***-------------|------------------------|----------------------------------------------------------***;
*** 16Sep2026   | Manivannan Mathialagan | Initial version created                                  ***;
***-------------------------------------------------------------------------------------------------***;

%macro mm_distinct(ds=, vars=, where=, out=_mm_distinct);
    proc sql;
        create table &out as
        select &vars, count(*) as count
        from &ds
        %if %length(%superq(where)) %then %do;
            where %unquote(&where)
        %end;
        group by &vars
        order by &vars;
    quit;

    title "MM_DISTINCT: &ds";
    proc print data=&out noobs;
    run;
    title;
%mend mm_distinct;
