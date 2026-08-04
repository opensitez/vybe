*> vybe-test: cobol/category_picture_clauses/test_pic_editing_decimal_point
*> origin: languages/cobol/tests/cobol/test_category_picture_clauses.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-DOT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 NUM PIC 9(3)V99 VALUE 123.45.
       01 EDITED PIC 999.99.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY EDITED.
    MOVE SPACES TO WS-VYBE-L
    STRING EDITED DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "123.45"
        DISPLAY "FAIL: want [123.45] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

