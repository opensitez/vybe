*> vybe-test: cobol/category_data_division/test_data_renames
*> origin: languages/cobol/tests/cobol/test_category_data_division.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-RENAME.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 REC.
          05 FLD-A PIC X VALUE "A".
          05 FLD-B PIC X VALUE "B".
          05 FLD-C PIC X VALUE "C".
       66 ALIAS-AC RENAMES FLD-A THRU FLD-C.
       PROCEDURE DIVISION.
           DISPLAY ALIAS-AC.
    MOVE SPACES TO WS-VYBE-L
    STRING ALIAS-AC DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABC"
        DISPLAY "FAIL: want [ABC] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

