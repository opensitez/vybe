*> vybe-test: cobol/condition_compound/condition_abbreviated_relation_and_range
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(2) VALUE 50.
PROCEDURE DIVISION.
    IF N >= 10 AND <= 90
        DISPLAY "IN RANGE"
    END-IF.
    STOP RUN.

