*> vybe-test: cobol/control_flow/perform_varying_down
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING I FROM 10 BY -1 UNTIL I < 1
        DISPLAY I
    END-PERFORM.
    STOP RUN.

