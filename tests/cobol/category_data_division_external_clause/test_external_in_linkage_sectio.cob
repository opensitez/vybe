*> vybe-test: cobol/category_data_division_external_clause/test_external_in_linkage_section
*> origin: languages/cobol/tests/cobol/test_category_data_division_external_clause.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. LINKAGE SECTION. 01 E PIC X VALUE 'L' IS EXTERNAL. PROCEDURE DIVISION USING E. DISPLAY E.
    MOVE SPACES TO WS-VYBE-L
    STRING E DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "L"
        DISPLAY "FAIL: want [L] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

