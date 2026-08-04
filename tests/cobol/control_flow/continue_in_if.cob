*> vybe-test: cobol/control_flow/continue_in_if
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9(3) VALUE 5.
PROCEDURE DIVISION.
    IF X > 10
        DISPLAY "Big"
    ELSE
        CONTINUE
    END-IF.
    STOP RUN.

