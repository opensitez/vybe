*> vybe-test: cobol/control_flow/perform_varying_step2
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING I FROM 0 BY 2 UNTIL I > 20
        DISPLAY I
    END-PERFORM.
    STOP RUN.

