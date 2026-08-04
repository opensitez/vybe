*> vybe-test: cobol/category_data_division_synchronized/test_sync_comp_sync
*> origin: languages/cobol/tests/cobol/test_category_data_division_synchronized.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC S9(4) COMP SYNCHRONIZED VALUE 123. PROCEDURE DIVISION. DISPLAY V.
    MOVE SPACES TO WS-VYBE-L
    STRING V DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0123"
        DISPLAY "FAIL: want [0123] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

