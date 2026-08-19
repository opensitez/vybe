*> vybe-test: cobol/arithmetic_control_flow/subtract_from_updates_target
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 20.
01 B PIC 9(3) VALUE 7.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    SUBTRACT B FROM A.
    DISPLAY A.
    MOVE SPACES TO WS-VYBE-L
    STRING A DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "013"
        DISPLAY "FAIL: want [013] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

