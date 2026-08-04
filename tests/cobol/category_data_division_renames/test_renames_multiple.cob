*> vybe-test: cobol/category_data_division_renames/test_renames_multiple
*> origin: languages/cobol/tests/cobol/test_category_data_division_renames.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 G. 05 A PIC X VALUE '1'. 05 B PIC X VALUE '2'. 66 R1 RENAMES A. 66 R2 RENAMES B. PROCEDURE DIVISION. DISPLAY R1 R2.
    MOVE SPACES TO WS-VYBE-L
    STRING R1 DELIMITED SIZE R2 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "12"
        DISPLAY "FAIL: want [12] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

