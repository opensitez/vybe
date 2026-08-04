*> vybe-test: cobol/category_data_division/test_data_value_clause_spaces
*> origin: languages/cobol/tests/cobol/test_category_data_division.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-VAL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 WS-A PIC X(3) VALUE SPACE.
       PROCEDURE DIVISION.
           DISPLAY "[" WS-A "]".
    MOVE SPACES TO WS-VYBE-L
    STRING "[" DELIMITED SIZE WS-A DELIMITED SIZE "]" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[   ]"
        DISPLAY "FAIL: want [[   ]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

