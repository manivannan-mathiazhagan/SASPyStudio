***-------------------------------------------------------------------------------------------------***;
*** Macro Name:    mm_freq.sas                                                                      ***;
*** Purpose:       Produce quick one-way or crossed frequency tables.                               ***;
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
*** VARS     | PROC FREQ TABLES expression                            | No default      | Yes       ***;
*** WHERE    | WHERE condition used to filter observations            | Blank           | No        ***;
*** MISSING  | Include missing values in frequency calculations       | Y               | No        ***;
*** LIST     | Display crossed tables in LIST format                  | N               | No        ***;
***-------------------------------------------------------------------------------------------------***;
*** Output(s):      PROC FREQ output. No datasets or macro variables are created.                   ***;
*** Dependencies:  Variables referenced in VARS= and WHERE= must exist in DS=.                      ***;
***-------------------------------------------------------------------------------------------------***;
*** Modification History                                                                            ***;
***-------------------------------------------------------------------------------------------------***;
*** Date        | Modified By            | Description                                              ***;
***-------------|------------------------|----------------------------------------------------------***;
*** 16Sep2026   | Manivannan Mathialagan | Initial version created                                  ***;
***-------------------------------------------------------------------------------------------------***;

%macro mm_freq(ds=, vars=, where=, missing=Y, list=N);

    title "MM_FREQ: &ds";

    proc freq data=&ds;

        %if %length(%superq(where)) %then %do;
            where %unquote(&where);
        %end;

        tables &vars
            %if %upcase(&missing)=Y or %upcase(&list)=Y %then %do;
                /
                %if %upcase(&missing)=Y %then missing;
                %if %upcase(&list)=Y %then list;
            %end;
        ;

    run;

    title;

%mend mm_freq;
