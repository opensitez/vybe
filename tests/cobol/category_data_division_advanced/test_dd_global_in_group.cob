*> vybe-test: cobol/category_data_division_advanced/test_dd_global_in_group
*> origin: languages/cobol/tests/cobol/test_category_data_division_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 G PIC X VALUE 'G' IS GLOBAL. 05 V PIC X VALUE 'V'. PROCEDURE DIVISION. DISPLAY G.
    MOVE SPACES TO WS-VYBE-L
    STRING G DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "G"
        DISPLAY "FAIL: want [G] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

