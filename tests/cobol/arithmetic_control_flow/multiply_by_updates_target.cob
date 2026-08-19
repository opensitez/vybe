*> vybe-test: cobol/arithmetic_control_flow/multiply_by_updates_target
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 6.
01 B PIC 9(3) VALUE 7.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MULTIPLY A BY B.
    DISPLAY B.
    MOVE SPACES TO WS-VYBE-L
    STRING B DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "042"
        DISPLAY "FAIL: want [042] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

