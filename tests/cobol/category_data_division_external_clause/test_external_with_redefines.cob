*> vybe-test: cobol/category_data_division_external_clause/test_external_with_redefines
*> origin: languages/cobol/tests/cobol/test_category_data_division_external_clause.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 E PIC X(4) VALUE 'ABCD' IS EXTERNAL. 01 R REDEFINES E PIC 9999. PROCEDURE DIVISION. DISPLAY E.
    MOVE SPACES TO WS-VYBE-L
    STRING E DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABCD"
        DISPLAY "FAIL: want [ABCD] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

