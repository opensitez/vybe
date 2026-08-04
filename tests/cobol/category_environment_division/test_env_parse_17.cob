*> vybe-test: cobol/category_environment_division/test_env_parse_17
*> origin: languages/cobol/tests/cobol/test_category_environment_division.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. CONFIGURATION SECTION. SPECIAL-NAMES. CURRENCY SIGN IS '$'. PROCEDURE DIVISION. DISPLAY 'ENV17'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'ENV17' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ENV17"
        DISPLAY "FAIL: want [ENV17] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

