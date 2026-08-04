*> vybe-test: cobol/conditions_extended/if_with_else_branch_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 3.
PROCEDURE DIVISION.
    IF WS-A > 5
        DISPLAY "BIG"
    ELSE
        DISPLAY "SMALL"
    END-IF.
    STOP RUN.

