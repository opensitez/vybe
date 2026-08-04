*> vybe-test: cobol/category_data_division_renames/test_renames_thru
*> origin: languages/cobol/tests/cobol/test_category_data_division_renames.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 G. 05 A PIC X VALUE '1'. 05 B PIC X VALUE '2'. 05 C PIC X VALUE '3'. 66 R RENAMES A THRU B. PROCEDURE DIVISION. DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "12"
        DISPLAY "FAIL: want [12] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

