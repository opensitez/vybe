*> vybe-test: cobol/category_picture_clauses/test_pic_alpha
*> origin: languages/cobol/tests/cobol/test_category_picture_clauses.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-A.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 VAL PIC A(5).
       PROCEDURE DIVISION.
           MOVE "ABCDE" TO VAL.
           DISPLAY VAL.
    MOVE SPACES TO WS-VYBE-L
    STRING VAL DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABCDE"
        DISPLAY "FAIL: want [ABCDE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

