***-------------------------------------------------------------------------------------------------***;
*** Macro Name:    mm_contents.sas                                                                  ***;
*** Purpose:       Display concise variable metadata for a SAS dataset.                             ***;
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
*** OUT      | Optional output metadata dataset                       | _MM_CONTENTS    | No        ***;
***-------------------------------------------------------------------------------------------------***;
*** Output(s):      Metadata dataset and PROC PRINT output.                                         ***;
*** Dependencies:  The dataset specified in DS= must exist.                                         ***;
***-------------------------------------------------------------------------------------------------***;
*** Modification History                                                                            ***;
***-------------------------------------------------------------------------------------------------***;
*** Date        | Modified By            | Description                                              ***;
***-------------|------------------------|----------------------------------------------------------***;
*** 16Sep2026   | Manivannan Mathialagan | Initial version created                                  ***;
***-------------------------------------------------------------------------------------------------***;

%macro mm_contents(ds=, out=_mm_contents);
    /* Capture metadata without printing the full PROC CONTENTS output. */
    proc contents data=&ds out=&out(keep=varnum name type length format informat label) noprint;
    run;

    proc sort data=&out;
        by varnum;
    run;

    title "MM_CONTENTS: &ds";
    proc print data=&out noobs label;
        var varnum name type length format informat label;
        label varnum='Position' name='Variable' type='Type' length='Length'
              format='Format' informat='Informat' label='Label';
    run;
    title;
%mend mm_contents;
