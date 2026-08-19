*> vybe-test: cobol/category_table_handling_advanced/test_tbl_runtime_reference_modification
*> origin: languages/cobol/tests/cobol/test_category_table_handling_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 TBL. 05 EL OCCURS 2 TIMES PIC X(4) VALUE 'ABCD'. PROCEDURE DIVISION. DISPLAY EL(2)(2:2).
    MOVE SPACES TO WS-VYBE-L
    STRING EL(2)(2:2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BC"
        DISPLAY "FAIL: want [BC] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

