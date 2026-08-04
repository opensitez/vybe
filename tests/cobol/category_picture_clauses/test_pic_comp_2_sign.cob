*> vybe-test: cobol/category_picture_clauses/test_pic_comp_2_sign
*> origin: languages/cobol/tests/cobol/test_category_picture_clauses.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-COMP.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 NUM PIC S9(3) COMP-2 VALUE -123.
       01 NUM2 PIC S9(3) COMP-2.
       PROCEDURE DIVISION.
           MOVE NUM TO NUM2.
           DISPLAY NUM2.
    MOVE SPACES TO WS-VYBE-L
    STRING NUM2 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "-123"
        DISPLAY "FAIL: want [-123] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

