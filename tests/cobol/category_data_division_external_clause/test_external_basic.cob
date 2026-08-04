*> vybe-test: cobol/category_data_division_external_clause/test_external_basic
*> origin: languages/cobol/tests/cobol/test_category_data_division_external_clause.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 E PIC X VALUE 'E' IS EXTERNAL. PROCEDURE DIVISION. DISPLAY E.
    MOVE SPACES TO WS-VYBE-L
    STRING E DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "E"
        DISPLAY "FAIL: want [E] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

