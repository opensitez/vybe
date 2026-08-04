*> vybe-test: cobol/category_data_division_advanced/test_dd_redefines_nested
*> origin: languages/cobol/tests/cobol/test_category_data_division_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 A PIC X(4) VALUE '1234'. 01 B REDEFINES A. 05 B1 PIC X(2). 05 B2 PIC X(2). PROCEDURE DIVISION. DISPLAY B2.
    MOVE SPACES TO WS-VYBE-L
    STRING B2 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "34"
        DISPLAY "FAIL: want [34] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

