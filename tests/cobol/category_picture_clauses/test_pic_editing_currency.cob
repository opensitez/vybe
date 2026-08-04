*> vybe-test: cobol/category_picture_clauses/test_pic_editing_currency
*> origin: languages/cobol/tests/cobol/test_category_picture_clauses.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-CURR.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 NUM PIC 9(4) VALUE 0123.
       01 EDITED PIC $$$,$$9.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY "[" EDITED "]".
    MOVE SPACES TO WS-VYBE-L
    STRING "[" DELIMITED SIZE EDITED DELIMITED SIZE "]" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[   $123]"
        DISPLAY "FAIL: want [[   $123]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

