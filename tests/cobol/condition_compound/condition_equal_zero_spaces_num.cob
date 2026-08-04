*> vybe-test: cobol/condition_compound/condition_equal_zero_spaces_numeric
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(3) VALUE ZEROS.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF N = ZEROS
        DISPLAY "ZERO"
    ELSE
        DISPLAY "NONZERO"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "ZERO" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ZERO"
        DISPLAY "FAIL: want [ZERO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

