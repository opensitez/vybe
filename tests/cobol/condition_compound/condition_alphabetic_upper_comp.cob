*> vybe-test: cobol/condition_compound/condition_alphabetic_upper_compiles
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(5) VALUE "HELLO".
PROCEDURE DIVISION.
    IF S IS ALPHABETIC-UPPER
        DISPLAY "UPPER"
    END-IF.
    STOP RUN.

