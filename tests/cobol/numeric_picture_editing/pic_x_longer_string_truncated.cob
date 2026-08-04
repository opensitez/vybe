*> vybe-test: cobol/numeric_picture_editing/pic_x_longer_string_truncated
*> origin: languages/cobol/tests/cobol/test_numeric_picture_editing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 DST PIC X(4) VALUE "XXXX".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "ABCDEFGH" TO DST.
    DISPLAY DST.
    MOVE SPACES TO WS-VYBE-L
    STRING DST DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABCD"
        DISPLAY "FAIL: want [ABCD] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

