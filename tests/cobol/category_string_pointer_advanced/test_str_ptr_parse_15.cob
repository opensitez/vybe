*> vybe-test: cobol/category_string_pointer_advanced/test_str_ptr_parse_15
*> origin: languages/cobol/tests/cobol/test_category_string_pointer_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION. DISPLAY 'OK'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'OK' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

