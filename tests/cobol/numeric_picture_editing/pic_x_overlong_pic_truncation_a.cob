*> vybe-test: cobol/numeric_picture_editing/pic_x_overlong_pic_truncation_at_definition_size
*> origin: languages/cobol/tests/cobol/test_numeric_picture_editing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(3) VALUE "AB".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY S.
    MOVE SPACES TO WS-VYBE-L
    STRING S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AB "
        DISPLAY "FAIL: want [AB ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

