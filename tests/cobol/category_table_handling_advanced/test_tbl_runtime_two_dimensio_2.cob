*> vybe-test: cobol/category_table_handling_advanced/test_tbl_runtime_two_dimensional_display
*> origin: languages/cobol/tests/cobol/test_category_table_handling_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 M. 05 R OCCURS 2 TIMES. 10 C OCCURS 3 TIMES PIC 9. PROCEDURE DIVISION. MOVE 7 TO C(2 3) DISPLAY C(2 3).
    MOVE SPACES TO WS-VYBE-L
    STRING C(2 DELIMITED SIZE 3) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "7"
        DISPLAY "FAIL: want [7] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

