*> vybe-test: cobol/category_picture_clauses/test_pic_numeric_implied_decimal
*> origin: languages/cobol/tests/cobol/test_category_picture_clauses.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-V.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 VAL PIC 9(3)V99 VALUE 123.45.
       PROCEDURE DIVISION.
           DISPLAY VAL.
    MOVE SPACES TO WS-VYBE-L
    STRING VAL DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "12345"
        DISPLAY "FAIL: want [12345] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

