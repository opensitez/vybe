*> vybe-test: cobol/category_data_division/test_data_66_set_condition_name
*> origin: languages/cobol/tests/cobol/test_category_data_division.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-66.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 WS-NUM PIC 9 VALUE 1.
       88 LOW VALUE 0 THRU 1.
       PROCEDURE DIVISION.
           IF LOW
               DISPLAY "LOW"
           END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "LOW" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "LOW"
        DISPLAY "FAIL: want [LOW] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

