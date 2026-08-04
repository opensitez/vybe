*> vybe-test: cobol/control_flow/if_and
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9(3) VALUE 5.
01 Y PIC 9(3) VALUE 10.
PROCEDURE DIVISION.
    IF X > 0 AND Y > 0
        DISPLAY "Both positive"
    END-IF.
    STOP RUN.

