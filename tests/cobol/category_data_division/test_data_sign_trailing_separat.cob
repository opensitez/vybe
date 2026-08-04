*> vybe-test: cobol/category_data_division/test_data_sign_trailing_separate
*> origin: languages/cobol/tests/cobol/test_category_data_division.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-SIGN-TR.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 VAL PIC S9(3) SIGN IS TRAILING SEPARATE CHARACTER.
       PROCEDURE DIVISION.
           MOVE -456 TO VAL.
           DISPLAY VAL.
    MOVE SPACES TO WS-VYBE-L
    STRING VAL DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "456-"
        DISPLAY "FAIL: want [456-] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

