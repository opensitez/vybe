*> vybe-test: cobol/subtract_advanced/test_subtract_multiple
*> origin: languages/cobol/tests/cobol/test_subtract_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9(3) VALUE 20.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    SUBTRACT 3 4 FROM WS-X.
    DISPLAY WS-X.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-X DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "013"
        DISPLAY "FAIL: want [013] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

