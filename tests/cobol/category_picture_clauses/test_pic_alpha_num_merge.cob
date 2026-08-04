*> vybe-test: cobol/category_picture_clauses/test_pic_alpha_num_merge
*> origin: languages/cobol/tests/cobol/test_category_picture_clauses.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-MIX.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 A PIC A(3) VALUE "A1 ".
       01 B PIC X(5) VALUE SPACES.
       PROCEDURE DIVISION.
           MOVE A TO B.
           DISPLAY B.
    MOVE SPACES TO WS-VYBE-L
    STRING B DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A1   "
        DISPLAY "FAIL: want [A1   ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

