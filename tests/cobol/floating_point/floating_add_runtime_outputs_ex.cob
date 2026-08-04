*> vybe-test: cobol/floating_point/floating_add_runtime_outputs_expected_value
*> origin: languages/cobol/tests/cobol/test_floating_point.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. FP11.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 A USAGE COMP-1 VALUE 1.
01 B USAGE COMP-1 VALUE 2.
01 C USAGE COMP-1.
PROCEDURE DIVISION.
    ADD B TO A
    MOVE A TO C
    DISPLAY C
    STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING C DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "3"
        DISPLAY "FAIL: want [3] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

