*> vybe-test: cobol/conditions_extended/if_numeric_class_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(3) VALUE "123".
PROCEDURE DIVISION.
    IF WS-A IS NUMERIC
        DISPLAY "NUM"
    END-IF.
    STOP RUN.

