*> vybe-test: cobol/control_flow/if_not_equal
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9(3) VALUE 5.
PROCEDURE DIVISION.
    IF X NOT = 0
        DISPLAY "Non-zero"
    END-IF.
    STOP RUN.

