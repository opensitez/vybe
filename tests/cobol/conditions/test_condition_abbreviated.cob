*> vybe-test: cobol/conditions/test_condition_abbreviated
*> origin: languages/cobol/tests/cobol/test_conditions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9 VALUE 5.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    IF WS-A > 0 AND < 10
        DISPLAY "BETWEEN"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "BETWEEN" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BETWEEN"
        DISPLAY "FAIL: want [BETWEEN] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

