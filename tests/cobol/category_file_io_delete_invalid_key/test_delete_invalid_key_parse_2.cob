*> vybe-test: cobol/category_file_io_delete_invalid_key/test_delete_invalid_key_parse_11
*> origin: languages/cobol/tests/cobol/test_category_file_io_delete_invalid_key.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'OK' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

