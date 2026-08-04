*> vybe-test: cobol/category_copybook_advanced/test_cp_copy_in_linkage_section
*> origin: languages/cobol/tests/cobol/test_category_copybook_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. LINKAGE SECTION. COPY 'MOCK'. PROCEDURE DIVISION. DISPLAY 'OK'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'OK' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

