*> vybe-test: cobol/structured_control_flow_expanded/perform_until_with_inner_if_compiles
*> origin: languages/cobol/tests/cobol/test_structured_control_flow_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM UNTIL WS-I >= 5
        ADD 1 TO WS-I
        IF WS-I = 3
            DISPLAY "MID"
        END-IF
    END-PERFORM.
    STOP RUN.

