*> vybe-test: cobol/category_picture_clauses/test_pic_numeric_zero_fill
*> origin: languages/cobol/tests/cobol/test_category_picture_clauses.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-ZFILL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 NUM PIC 9999 VALUE 12.
       PROCEDURE DIVISION.
           MOVE NUM TO NUM.
           DISPLAY NUM.
    MOVE SPACES TO WS-VYBE-L
    STRING NUM DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0012"
        DISPLAY "FAIL: want [0012] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

