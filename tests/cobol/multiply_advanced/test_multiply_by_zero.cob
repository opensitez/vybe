*> vybe-test: cobol/multiply_advanced/test_multiply_by_zero
*> origin: languages/cobol/tests/cobol/test_multiply_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 5.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MULTIPLY 0 BY WS-A.
    DISPLAY WS-A.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-A DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "000"
        DISPLAY "FAIL: want [000] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

