*> vybe-test: cobol/numeric_picture_editing/pic_9_to_9v99_extends_decimal
*> origin: languages/cobol/tests/cobol/test_numeric_picture_editing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC 9(3) VALUE 123.
01 DST PIC 9(3)V99 VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE SRC TO DST.
    DISPLAY DST.
    MOVE SPACES TO WS-VYBE-L
    STRING DST DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "12300"
        DISPLAY "FAIL: want [12300] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

