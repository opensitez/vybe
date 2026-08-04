*> vybe-test: cobol/value_all_forms/value_signed_negative_literal
*> origin: languages/cobol/tests/cobol/test_value_all_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC S9(4) VALUE -9.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY N.
    MOVE SPACES TO WS-VYBE-L
    STRING N DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "-0009"
        DISPLAY "FAIL: want [-0009] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

