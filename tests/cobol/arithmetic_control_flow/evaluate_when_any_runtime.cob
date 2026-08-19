*> vybe-test: cobol/arithmetic_control_flow/evaluate_when_any_runtime
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 5.
01 B PIC 9 VALUE 1.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE A ALSO B
        WHEN 1 ALSO ANY DISPLAY "NO"
        WHEN 5 ALSO 1
            DISPLAY "YES"
        WHEN OTHER
            DISPLAY "NA"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "NO" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "NO"
        DISPLAY "FAIL: want [NO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

