*> vybe-test: cobol/arithmetic_control_flow_matrix/evaluate_numeric_when_thru_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 7.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE X
        WHEN 1 THRU 5 DISPLAY "L"
        WHEN 6 THRU 9 DISPLAY "H"
        WHEN OTHER DISPLAY "O"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "L" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "H"
        DISPLAY "FAIL: want [H] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

