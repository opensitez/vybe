*> vybe-test: cobol/category_data_division_synchronized/test_sync_basic_right
*> origin: languages/cobol/tests/cobol/test_category_data_division_synchronized.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X(5) SYNCHRONIZED RIGHT VALUE 'A'. PROCEDURE DIVISION. DISPLAY V.
    MOVE SPACES TO WS-VYBE-L
    STRING V DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A    "
        DISPLAY "FAIL: want [A    ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

