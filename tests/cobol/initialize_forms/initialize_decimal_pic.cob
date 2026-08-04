*> vybe-test: cobol/initialize_forms/initialize_decimal_pic
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC 9(3)V99 VALUE 123.45.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INITIALIZE D.
    DISPLAY D.
    MOVE SPACES TO WS-VYBE-L
    STRING D DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "00000"
        DISPLAY "FAIL: want [00000] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

