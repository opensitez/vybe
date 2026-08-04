*> vybe-test: cobol/category_picture_clauses/test_pic_editing_comma
*> origin: languages/cobol/tests/cobol/test_category_picture_clauses.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-COMMA.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 NUM PIC 9(4) VALUE 1234.
       01 EDITED PIC 9,999.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY EDITED.
    MOVE SPACES TO WS-VYBE-L
    STRING EDITED DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1,234"
        DISPLAY "FAIL: want [1,234] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

