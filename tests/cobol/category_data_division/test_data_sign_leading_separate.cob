*> vybe-test: cobol/category_data_division/test_data_sign_leading_separate
*> origin: languages/cobol/tests/cobol/test_category_data_division.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-SIGN.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 VAL PIC S9(3) SIGN IS LEADING SEPARATE CHARACTER.
       PROCEDURE DIVISION.
           MOVE -123 TO VAL.
           DISPLAY VAL.
    MOVE SPACES TO WS-VYBE-L
    STRING VAL DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "-123"
        DISPLAY "FAIL: want [-123] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

