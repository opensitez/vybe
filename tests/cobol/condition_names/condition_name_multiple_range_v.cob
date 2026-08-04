*> vybe-test: cobol/condition_names/condition_name_multiple_range_values_compile
*> origin: languages/cobol/tests/cobol/test_condition_names.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-AGE PIC 99 VALUE 18.
   88 AGE-STATE VALUE 0 THRU 17, 18 THRU 30.
PROCEDURE DIVISION.

    IF AGE-STATE
        CONTINUE
    END-IF.
    STOP RUN.

