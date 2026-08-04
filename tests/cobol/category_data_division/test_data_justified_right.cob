*> vybe-test: cobol/category_data_division/test_data_justified_right
*> origin: languages/cobol/tests/cobol/test_category_data_division.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-JUST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 VAL PIC X(5) JUSTIFIED RIGHT.
       PROCEDURE DIVISION.
           MOVE "AB" TO VAL.
           DISPLAY "[" VAL "]".
    MOVE SPACES TO WS-VYBE-L
    STRING "[" DELIMITED SIZE VAL DELIMITED SIZE "]" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[   AB]"
        DISPLAY "FAIL: want [[   AB]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

