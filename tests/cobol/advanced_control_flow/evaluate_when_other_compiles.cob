*> vybe-test: cobol/advanced_control_flow/evaluate_when_other_compiles
*> origin: languages/cobol/tests/cobol/test_advanced_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(1) VALUE 9.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE WS-A
        WHEN 1
            DISPLAY "ONE"
        WHEN 2
            DISPLAY "TWO"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "ONE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ONE"
        DISPLAY "FAIL: want [ONE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

