***-------------------------------------------------------------------------------------------------***;
*** Macro Name:    mm_check.sas                                                                     ***;
*** Purpose:       Run a compact combined health check of a SAS dataset.                            ***;
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
*** KEY      | Variables expected to uniquely identify observations   | Blank           | No        ***;
*** VARS     | Variables used for missing-value and frequency checks  | Blank           | No        ***;
*** WHERE    | WHERE condition applied to checks                      | Blank           | No        ***;
*** FREQ     | Run frequency tables for VARS=                         | Y               | No        ***;
*** CONTENTS | Display concise variable metadata                      | N               | No        ***;
***-------------------------------------------------------------------------------------------------***;
*** Output(s):      Dataset summary in log and requested diagnostic output. Temporary WORK datasets ***;
***                 _MM_DUPS, _MM_MISSING, and _MM_CONTENTS may be created.                         ***;
*** Dependencies:  Calls MM_DSINFO, MM_DUPS, MM_MISSING, MM_FREQ, and MM_CONTENTS.                  ***;
***-------------------------------------------------------------------------------------------------***;
*** Modification History                                                                            ***;
***-------------------------------------------------------------------------------------------------***;
*** Date        | Modified By            | Description                                              ***;
***-------------|------------------------|----------------------------------------------------------***;
*** 16Sep2026   | Manivannan Mathialagan | Initial version created                                  ***;
***-------------------------------------------------------------------------------------------------***;

%macro mm_check(ds=, key=, vars=, where=, freq=Y, contents=N);
    %put NOTE: ======================== MM_CHECK START ========================;
    %mm_dsinfo(ds=&ds);

    /* Check key uniqueness only when a key is supplied. */
    %if %length(%superq(key)) %then %do;
        %mm_dups(ds=&ds, by=&key, where=&where, out=_mm_dups, print=Y);
    %end;

    /* Review missingness and distributions for selected variables. */
    %if %length(%superq(vars)) %then %do;
        %mm_missing(ds=&ds, vars=&vars, where=&where, out=_mm_missing);
        %if %upcase(&freq)=Y %then %do;
            %mm_freq(ds=&ds, vars=&vars, where=&where, missing=Y);
        %end;
    %end;

    %if %upcase(&contents)=Y %then %do;
        %mm_contents(ds=&ds, out=_mm_contents);
    %end;
    %put NOTE: ========================= MM_CHECK END =========================;
%mend mm_check;
