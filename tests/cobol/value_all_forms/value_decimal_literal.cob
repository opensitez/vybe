*> vybe-test: cobol/value_all_forms/value_decimal_literal
*> origin: languages/cobol/tests/cobol/test_value_all_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC 9(3)V99 VALUE 12.34.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY D.
    MOVE SPACES TO WS-VYBE-L
    STRING D DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "01234"
        DISPLAY "FAIL: want [01234] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

