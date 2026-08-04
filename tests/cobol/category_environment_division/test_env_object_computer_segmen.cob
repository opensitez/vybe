*> vybe-test: cobol/category_environment_division/test_env_object_computer_segment_limit
*> origin: languages/cobol/tests/cobol/test_category_environment_division.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. CONFIGURATION SECTION. OBJECT-COMPUTER. X SEGMENT-LIMIT IS 50. PROCEDURE DIVISION. DISPLAY 'OK'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'OK' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

