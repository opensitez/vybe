*> vybe-test: cobol/figurative_constants_extended/figurative_constant_zeros_assign_numeric_value
*> origin: languages/cobol/tests/cobol/test_figurative_constants_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-N PIC 9(3) VALUE 123.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE ZEROS TO WS-N.
    DISPLAY WS-N.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-N DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "000"
        DISPLAY "FAIL: want [000] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

