***-------------------------------------------------------------------------------------------------***;
*** Macro Name:    mm_compare.sas                                                                   ***;
*** Purpose:       Compare two SAS datasets and optionally create a difference dataset.             ***;
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
*** BASE     | Base/reference SAS dataset                             | No default      | Yes       ***;
*** COMP     | Comparison SAS dataset                                 | No default      | Yes       ***;
*** ID       | Variables identifying corresponding observations       | Blank           | No        ***;
*** VAR      | Variables to compare blank compares common variables   | Blank           | No        ***;
*** CRITERION| Numeric comparison criterion                           | 1E-12           | No        ***;
*** METHOD   | PROC COMPARE numeric comparison method                 | ABSOLUTE        | No        ***;
*** OUT      | Optional output dataset containing differences         | _MM_COMPARE     | No        ***;
***-------------------------------------------------------------------------------------------------***;
*** Output(s):      PROC COMPARE output and the dataset specified in OUT=.                          ***;
*** Dependencies:  BASE= and COMP= datasets must exist ID= datasets should be sorted by ID=.        ***;
***-------------------------------------------------------------------------------------------------***;
*** Modification History                                                                            ***;
***-------------------------------------------------------------------------------------------------***;
*** Date        | Modified By            | Description                                              ***;
***-------------|------------------------|----------------------------------------------------------***;
*** 16Sep2026   | Manivannan Mathialagan | Initial version created                                  ***;
***-------------------------------------------------------------------------------------------------***;

%macro mm_compare(base=, comp=, id=, var=, criterion=1E-12,
                  method=absolute, out=_mm_compare);
    title "MM_COMPARE: &base versus &comp";
    proc compare base=&base compare=&comp
                 criterion=&criterion method=&method
                 out=&out outbase outcomp outdif outnoequal;
        %if %length(%superq(id)) %then %do;
            id &id;
        %end;
        %if %length(%superq(var)) %then %do;
            var &var;
        %end;
    run;
    title;
%mend mm_compare;
