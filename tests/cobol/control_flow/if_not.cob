*> vybe-test: cobol/control_flow/if_not
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    IF NOT X > 0
        DISPLAY "Zero or negative"
    END-IF.
    STOP RUN.

