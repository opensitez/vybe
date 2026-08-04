*> vybe-test: cobol/category_data_division_synchronized/test_sync_numeric_move_runtime
*> origin: languages/cobol/tests/cobol/test_category_data_division_synchronized.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 N PIC S9(4) SYNCHRONIZED VALUE -7. PROCEDURE DIVISION. DISPLAY N.
    MOVE SPACES TO WS-VYBE-L
    STRING N DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "-007"
        DISPLAY "FAIL: want [-007] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

