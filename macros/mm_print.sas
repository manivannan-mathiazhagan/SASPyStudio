***-------------------------------------------------------------------------------------------------***;
*** Macro Name:    mm_print.sas                                                                     ***;
*** Purpose:       Print a SAS dataset with optional variable selection, filtering, and row limit.  ***;
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
*** VARS     | Variables to display blank displays all variables      | Blank           | No        ***;
*** WHERE    | WHERE condition used to filter observations            | Blank           | No        ***;
*** N        | Maximum qualifying observations to display             | Blank (all)     | No        ***;
***-------------------------------------------------------------------------------------------------***;
*** Output(s):      PROC PRINT output. No permanent datasets or macro variables are created.        ***;
*** Dependencies:  The dataset specified in DS= must exist.                                         ***;
***-------------------------------------------------------------------------------------------------***;
*** Modification History                                                                            ***;
***-------------------------------------------------------------------------------------------------***;
*** Date        | Modified By            | Description                                              ***;
***-------------|------------------------|----------------------------------------------------------***;
*** 16Sep2026   | Manivannan Mathialagan | Initial version created                                  ***;
***-------------------------------------------------------------------------------------------------***;

%macro mm_print(ds=, vars=, where=, n=);
    %local _mm_dsopt;

    /* Build dataset options only when WHERE= or N= is supplied. */
    %let _mm_dsopt=;
    %if %length(%superq(where)) %then %let _mm_dsopt=&_mm_dsopt where=(%unquote(&where));
    %if %length(%superq(n))     %then %let _mm_dsopt=&_mm_dsopt obs=&n;

    title "MM_PRINT: &ds";
    proc print data=&ds %if %length(%superq(_mm_dsopt)) %then %do; (&_mm_dsopt) %end; noobs;
        /* Omit VAR statement to display every variable. */
        %if %length(%superq(vars)) %then %do;
            var &vars;
        %end;
    run;
    title;
%mend mm_print;
