*> vybe-test: cobol/condition_compound/condition_abbreviated_relation_or
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 5.
PROCEDURE DIVISION.
    IF N = 3 OR 5 OR 7
        DISPLAY "ODD"
    END-IF.
    STOP RUN.

