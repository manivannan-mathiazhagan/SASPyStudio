***-------------------------------------------------------------------------------------------------***;
*** Macro Name:    mm_dups.sas                                                                      ***;
*** Purpose:       Identify all observations belonging to duplicate key groups.                     ***;
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
*** BY       | Variables expected to uniquely identify observations   | No default      | Yes       ***;
*** WHERE    | WHERE condition applied before duplicate checking      | Blank           | No        ***;
*** OUT      | Output dataset containing every duplicate record       | _MM_DUPS        | No        ***;
*** PRINT    | Print duplicate observations                           | Y               | No        ***;
***-------------------------------------------------------------------------------------------------***;
*** Output(s):      Dataset specified in OUT= _MM_DUPKEYS is also created in WORK.                  ***;
*** Dependencies:  BY= variables must exist in DS=.                                                 ***;
***-------------------------------------------------------------------------------------------------***;
*** Modification History                                                                            ***;
***-------------------------------------------------------------------------------------------------***;
*** Date        | Modified By            | Description                                              ***;
***-------------|------------------------|----------------------------------------------------------***;
*** 16Sep2026   | Manivannan Mathialagan | Initial version created                                  ***;
***-------------------------------------------------------------------------------------------------***;

%macro mm_dups(ds=, by=, where=, out=_mm_dups, print=Y);
    /* Find key combinations occurring more than once. */
    proc sql;
        create table _mm_dupkeys as
        select &by, count(*) as _mm_count
        from &ds
        %if %length(%superq(where)) %then %do;
            where %unquote(&where)
        %end;
        group by &by
        having calculated _mm_count > 1;
    quit;

    /* Merge the duplicate keys back to retain every record in each duplicate group. */
    proc sort data=_mm_dupkeys;
        by &by;
    run;
    proc sort data=&ds
        %if %length(%superq(where)) %then %do; (where=(%unquote(&where))) %end;
        out=_mm_dups_source;
        by &by;
    run;
    data &out;
        merge _mm_dups_source(in=_mm_in) _mm_dupkeys(in=_mm_dup keep=&by);
        by &by;
        if _mm_in and _mm_dup;
        drop _mm_count;
    run;

    %if %upcase(&print)=Y %then %do;
        title "MM_DUPS: &ds by &by";
        proc print data=&out noobs;
        run;
        title;
    %end;
%mend mm_dups;
