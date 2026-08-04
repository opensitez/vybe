*> vybe-test: cobol/category_data_division_occurs/test_occurs_index_name_shadowing
*> origin: languages/cobol/tests/cobol/test_category_data_division_occurs.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 I PIC 9 VALUE 1. 01 TBL. 05 EL OCCURS 3 TIMES INDEXED BY I PIC 9 VALUE 2. PROCEDURE DIVISION. SET I TO 3. DISPLAY EL(I).
    MOVE SPACES TO WS-VYBE-L
    STRING EL(I) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2"
        DISPLAY "FAIL: want [2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

