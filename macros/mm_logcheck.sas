***-------------------------------------------------------------------------------------------------***;
*** Macro Name:    mm_logcheck.sas                                                                  ***;
*** Purpose:       Search a saved SAS log for common errors, warnings, and notable messages.        ***;
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
*** LOG      | Full path to the saved SAS log file                    | No default      | Yes       ***;
*** OUT      | Output dataset containing matching log lines           | _MM_LOGCHECK    | No        ***;
*** PRINT    | Print matching log lines                               | Y               | No        ***;
*** NOTES    | Include selected problematic NOTE messages             | Y               | No        ***;
***-------------------------------------------------------------------------------------------------***;
*** Output(s):      Dataset specified in OUT= and optional PROC PRINT output.                       ***;
*** Dependencies:  LOG= must reference a readable external log file.                                ***;
***-------------------------------------------------------------------------------------------------***;
*** Modification History                                                                            ***;
***-------------------------------------------------------------------------------------------------***;
*** Date        | Modified By            | Description                                              ***;
***-------------|------------------------|----------------------------------------------------------***;
*** 16Sep2026   | Manivannan Mathialagan | Initial version created                                  ***;
***-------------------------------------------------------------------------------------------------***;

%macro mm_logcheck(log=, out=_mm_logcheck, print=Y, notes=Y);
    data &out;
        length line_number 8 message $32767 category $7;
        infile "&log" lrecl=32767 truncover;
        input message $char32767.;
        line_number=_n_;

        /* Classify only messages that generally require programmer review. */
        if prxmatch('/^\s*ERROR[: ]/i',message) then category='ERROR';
        else if prxmatch('/^\s*WARNING[: ]/i',message) then category='WARNING';
        %if %upcase(&notes)=Y %then %do;
            else if prxmatch('/^\s*NOTE:.*(uninitialized|invalid|missing values were generated|repeats of by values|division by zero|stopped due to looping)/i',message)
                then category='NOTE';
        %end;
        if not missing(category);
    run;

    %if %upcase(&print)=Y %then %do;
        title "MM_LOGCHECK: &log";
        proc print data=&out noobs;
            var category line_number message;
        run;
        title;
    %end;
%mend mm_logcheck;
