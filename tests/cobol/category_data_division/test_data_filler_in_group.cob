*> vybe-test: cobol/category_data_division/test_data_filler_in_group
*> origin: languages/cobol/tests/cobol/test_category_data_division.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-FILL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 WS-REC.
          05 FILLER PIC X(3) VALUE "ABC".
          05 WS-FLD PIC X(3) VALUE "DEF".
       PROCEDURE DIVISION.
           DISPLAY WS-REC.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-REC DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABCDEF"
        DISPLAY "FAIL: want [ABCDEF] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

