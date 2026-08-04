*> vybe-test: cobol/numeric_picture_editing/pic_99_v_9999_decimal_leading_zero
*> origin: languages/cobol/tests/cobol/test_numeric_picture_editing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC 99V9999 VALUE 01.2345.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY D.
    MOVE SPACES TO WS-VYBE-L
    STRING D DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "012345"
        DISPLAY "FAIL: want [012345] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

