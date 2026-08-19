*> vybe-test: cobol/numeric_functions/intrinsic_length_of_literal
*> origin: languages/cobol/tests/cobol/test_numeric_functions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 L PIC 9(4) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    COMPUTE L = FUNCTION LENGTH("HELLO").
    DISPLAY L.
    MOVE SPACES TO WS-VYBE-L
    STRING L DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0005"
        DISPLAY "FAIL: want [0005] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

