*> vybe-test: cobol/numeric_picture_editing/pic_x_pads_right_when_shorter_source
*> origin: languages/cobol/tests/cobol/test_numeric_picture_editing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(3) VALUE "ABC".
01 DST PIC X(6) VALUE "XXXXXX".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE SRC TO DST.
    DISPLAY DST.
    MOVE SPACES TO WS-VYBE-L
    STRING DST DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABC   "
        DISPLAY "FAIL: want [ABC   ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

