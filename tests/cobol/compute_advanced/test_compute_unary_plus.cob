*> vybe-test: cobol/compute_advanced/test_compute_unary_plus
*> origin: languages/cobol/tests/cobol/test_compute_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC S9(3) VALUE 42.
01 WS-B PIC S9(3) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    COMPUTE WS-B = + WS-A.
    DISPLAY WS-B.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-B DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "+42"
        DISPLAY "FAIL: want [+42] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

