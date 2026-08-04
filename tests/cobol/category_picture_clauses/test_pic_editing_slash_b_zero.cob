*> vybe-test: cobol/category_picture_clauses/test_pic_editing_slash_b_zero
*> origin: languages/cobol/tests/cobol/test_category_picture_clauses.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-INS.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 NUM PIC 9(6) VALUE 123456.
       01 EDITED PIC 99/99/99.
       01 NUM2 PIC 9(2) VALUE 12.
       01 EDITED2 PIC 9B9.
       01 EDITED3 PIC 909.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           MOVE NUM2 TO EDITED2.
           MOVE NUM2 TO EDITED3.
           DISPLAY EDITED " " EDITED2 " " EDITED3.
    MOVE SPACES TO WS-VYBE-L
    STRING EDITED DELIMITED SIZE " " DELIMITED SIZE EDITED2 DELIMITED SIZE " " DELIMITED SIZE EDITED3 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "12/34/56 1 2 102"
        DISPLAY "FAIL: want [12/34/56 1 2 102] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

