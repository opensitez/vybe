*> vybe-test: cobol/subtract_advanced/test_subtract_giving_zero
*> origin: languages/cobol/tests/cobol/test_subtract_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 10.
01 WS-B PIC 9(3) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    SUBTRACT WS-A FROM WS-A GIVING WS-B.
    DISPLAY WS-B.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-B DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "000"
        DISPLAY "FAIL: want [000] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

