*> vybe-test: cobol/add_advanced/test_add_negative_literal
*> origin: languages/cobol/tests/cobol/test_add_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC S9(3) VALUE 10.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    ADD -5 TO WS-A.
    DISPLAY WS-A.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-A DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "005"
        DISPLAY "FAIL: want [005] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

