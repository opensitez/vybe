*> vybe-test: cobol/category_data_division_external_clause/test_external_signed_numeric
*> origin: languages/cobol/tests/cobol/test_category_data_division_external_clause.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 E-S NUM PIC S9(4) VALUE -12 IS EXTERNAL. PROCEDURE DIVISION. DISPLAY E-S.
    MOVE SPACES TO WS-VYBE-L
    STRING E-S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "-12"
        DISPLAY "FAIL: want [-12] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

