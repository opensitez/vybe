*> vybe-test: cobol/control_flow/perform_until_simple
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    PERFORM UNTIL I >= 10
        ADD 1 TO I
    END-PERFORM.
    STOP RUN.

