*> vybe-test: cobol/category_table_handling_advanced/test_tbl_runtime_two_dimensional_sum
*> origin: languages/cobol/tests/cobol/test_category_table_handling_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 M. 05 R OCCURS 2 TIMES. 10 C OCCURS 2 TIMES PIC 9. 01 S PIC 99 VALUE 0. PROCEDURE DIVISION. MOVE 1 TO C(1 1) MOVE 4 TO C(2 2) ADD C(1 1) TO S ADD C(2 2) TO S DISPLAY S STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "5"
        DISPLAY "FAIL: want [5] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

