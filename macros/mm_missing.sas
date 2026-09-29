***-------------------------------------------------------------------------------------------------***;
*** Macro Name:    mm_missing.sas                                                                   ***;
*** Purpose:       Summarize missing and nonmissing values for selected variables.                  ***;
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
*** VARS     | Space-separated variables to evaluate                  | No default      | Yes       ***;
*** WHERE    | WHERE condition used to filter observations            | Blank           | No        ***;
*** OUT      | Output summary dataset                                 | _MM_MISSING     | No        ***;
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

%macro mm_missing(ds=, vars=, where=, out=_mm_missing);
    %local _mm_i _mm_var _mm_nvars;
    %let _mm_nvars=%sysfunc(countw(%superq(vars),%str( )));

    data &out;
        length variable $32 total missing nonmissing 8 percent_missing 8;
        stop;
    run;

    %do _mm_i=1 %to &_mm_nvars;
        %let _mm_var=%scan(%superq(vars),&_mm_i,%str( ));
        proc sql;
            insert into &out
            select "&_mm_var", count(*),
                   sum(missing(&_mm_var)),
                   sum(not missing(&_mm_var)),
                   case when count(*) > 0 then 100*sum(missing(&_mm_var))/count(*) else . end
            from &ds
            %if %length(%superq(where)) %then %do;
                where %unquote(&where)
            %end;
            ;
        quit;
    %end;

    title "MM_MISSING: &ds";
    proc print data=&out noobs;
        format percent_missing 6.2;
    run;
    title;
%mend mm_missing;
