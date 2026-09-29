/*=============================================================
  SASPy Studio - Second Program Test
=============================================================*/

%put NOTE: ================================================;
%put NOTE: Second SASPy Studio program started.;
%put NOTE: ================================================;

data work.class_test;
    set sashelp.class;
    BMI = (weight / (height * height)) * 703;
run;

%put NOTE: WORK.CLASS_TEST created.;

title "SASPy Studio - SASHELP.CLASS Test";

proc print data=work.class_test;
    var name sex age height weight BMI;
    format BMI 8.2;
run;

title;

proc means data=work.class_test n mean std min max maxdec=2;
    var age height weight BMI;
run;

filename csvout "%sysfunc(pathname(work))/class_test.csv";

proc export data=work.class_test
    outfile=csvout
    dbms=csv
    replace;
run;

filename csvout clear;

%put NOTE: CLASS_TEST.CSV created.;
%put NOTE: Second SASPy Studio program completed.;