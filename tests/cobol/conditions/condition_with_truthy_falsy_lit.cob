*> vybe-test: cobol/conditions/condition_with_truthy_falsy_literals_compile
*> origin: languages/cobol/tests/cobol/test_conditions.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 FLAG PIC X VALUE "Y".
PROCEDURE DIVISION.
    IF FLAG = "Y"
        DISPLAY "YES"
    ELSE
        DISPLAY "NO"
    END-IF.
    STOP RUN.

