*> vybe-test: cobol/nested_if_else/if_equal_zeros_check
*> origin: languages/cobol/tests/cobol/test_nested_if_else.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(5) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF N = ZEROS
        DISPLAY "ALL ZERO"
    ELSE
        DISPLAY "NONZERO"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "ALL ZERO" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ALL ZERO"
        DISPLAY "FAIL: want [ALL ZERO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

