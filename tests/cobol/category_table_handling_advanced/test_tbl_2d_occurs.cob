*> vybe-test: cobol/category_table_handling_advanced/test_tbl_2d_occurs
*> origin: languages/cobol/tests/cobol/test_category_table_handling_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 TBL. 05 R OCCURS 2 TIMES. 10 C OCCURS 2 TIMES PIC 9 VALUE 2. PROCEDURE DIVISION. DISPLAY C(2 2).
    MOVE SPACES TO WS-VYBE-L
    STRING C(2 DELIMITED SIZE 2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2"
        DISPLAY "FAIL: want [2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

