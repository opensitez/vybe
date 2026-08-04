*> vybe-test: cobol/category_picture_clauses/test_pic_editing_cr_db
*> origin: languages/cobol/tests/cobol/test_category_picture_clauses.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-CRDB.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 NUM1 PIC S9(3) VALUE -123.
       01 NUM2 PIC S9(3) VALUE 123.
       01 EDITED1 PIC 999CR.
       01 EDITED2 PIC 999DB.
       PROCEDURE DIVISION.
           MOVE NUM1 TO EDITED1.
           MOVE NUM2 TO EDITED2.
           DISPLAY EDITED1 " " EDITED2.
    MOVE SPACES TO WS-VYBE-L
    STRING EDITED1 DELIMITED SIZE " " DELIMITED SIZE EDITED2 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "123CR 123  "
        DISPLAY "FAIL: want [123CR 123  ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

