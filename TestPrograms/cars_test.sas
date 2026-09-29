
/*=============================================================
  SASPy Studio - End-to-End Test
=============================================================*/

%put NOTE: ================================================;
%put NOTE: SASPy Studio test started.;
%put NOTE: Testing LOG, RESULTS and OUTPUT.;
%put NOTE: ================================================;

/*-------------------------------------------------------------
  1. Create WORK dataset from SASHELP.CARS
-------------------------------------------------------------*/

data work.cars_test;
    set sashelp.cars;

    where type in ("SUV", "Sedan");

    Price_Difference = Invoice - MSRP;

    label
        Price_Difference = "Invoice minus MSRP";
run;

%put NOTE: WORK.CARS_TEST has been created.;


/*-------------------------------------------------------------
  2. Put useful information into SAS LOG
-------------------------------------------------------------*/

proc sql noprint;
    select count(*)
        into :car_count trimmed
    from work.cars_test;
quit;

%put NOTE: Number of selected cars = &car_count.;


/*-------------------------------------------------------------
  3. HTML RESULTS
-------------------------------------------------------------*/

title "SASPy Studio Test - Selected Cars";

proc print data=work.cars_test(obs=15) label;
    var Make
        Model
        Type
        Origin
        MSRP
        Invoice
        Price_Difference;
run;


title "SASPy Studio Test - Price Summary";

proc means data=work.cars_test
    n
    mean
    std
    min
    max
    maxdec=2;

    class Type;
    var MSRP Invoice;
run;

title;


/*-------------------------------------------------------------
  4. OUTPUT FILE

  TEMPORARY SAS-server location.
  SASPy Studio will later download this into its output folder.
-------------------------------------------------------------*/

filename csvout "%sysfunc(pathname(work))/cars_test.csv";

proc export
    data=work.cars_test
    outfile=csvout
    dbms=csv
    replace;
run;

%put NOTE: CSV output created on SAS server.;
%put NOTE: SASPy Studio test completed.;