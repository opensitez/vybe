*> vybe-test: cobol/category_data_division_renames/test_renames_mixed_numeric_and_alpha
*> origin: languages/cobol/tests/cobol/test_category_data_division_renames.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R. 05 N PIC 99 VALUE 12. 05 T PIC X(2) VALUE 'AB'. 66 ALIAS RENAMES N. PROCEDURE DIVISION. DISPLAY ALIAS.
    MOVE SPACES TO WS-VYBE-L
    STRING ALIAS DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "12"
        DISPLAY "FAIL: want [12] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

