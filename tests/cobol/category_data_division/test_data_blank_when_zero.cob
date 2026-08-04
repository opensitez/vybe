*> vybe-test: cobol/category_data_division/test_data_blank_when_zero
*> origin: languages/cobol/tests/cobol/test_category_data_division.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-BWZ.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 VAL PIC 9(3) BLANK WHEN ZERO.
       PROCEDURE DIVISION.
           MOVE 0 TO VAL.
           DISPLAY "[" VAL "]".
    MOVE SPACES TO WS-VYBE-L
    STRING "[" DELIMITED SIZE VAL DELIMITED SIZE "]" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[   ]"
        DISPLAY "FAIL: want [[   ]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

