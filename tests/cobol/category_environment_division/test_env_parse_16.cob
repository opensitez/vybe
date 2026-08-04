*> vybe-test: cobol/category_environment_division/test_env_parse_16
*> origin: languages/cobol/tests/cobol/test_category_environment_division.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. CONFIGURATION SECTION. SPECIAL-NAMES. DECIMAL-POINT IS COMMA. PROCEDURE DIVISION. DISPLAY 'ENV16'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'ENV16' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ENV16"
        DISPLAY "FAIL: want [ENV16] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

