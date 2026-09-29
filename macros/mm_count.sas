***-------------------------------------------------------------------------------------------------***;
*** Macro Name:    mm_count.sas                                                                     ***;
*** Purpose:       Count observations satisfying an optional condition and write the count to log.  ***;
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
*** WHERE    | WHERE condition used to filter observations            | Blank           | No        ***;
*** OUTVAR   | Macro variable that receives the count                 | MM_COUNT        | No        ***;
***-------------------------------------------------------------------------------------------------***;
*** Output(s):      Creates the global macro variable named by OUTVAR= and writes its value to log. ***;
*** Dependencies:  The dataset specified in DS= must exist.                                         ***;
***-------------------------------------------------------------------------------------------------***;
*** Modification History                                                                            ***;
***-------------------------------------------------------------------------------------------------***;
*** Date        | Modified By            | Description                                              ***;
***-------------|------------------------|----------------------------------------------------------***;
*** 16Sep2026   | Manivannan Mathialagan | Initial version created                                  ***;
***-------------------------------------------------------------------------------------------------***;

%macro mm_count(ds=, where=, outvar=MM_COUNT);
    %global &outvar;
    proc sql noprint;
        select count(*) into :&outvar trimmed
        from &ds
        %if %length(%superq(where)) %then %do;
            where %unquote(&where)
        %end;
        ;
    quit;
    %put NOTE: [MM_COUNT] Dataset=&ds Count=&&&outvar;
%mend mm_count;
