*> vybe-test: cobol/category_environment_division/test_env_parse_11
*> origin: languages/cobol/tests/cobol/test_category_environment_division.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. CONFIGURATION SECTION. SOURCE-COMPUTER. IBM-37. PROCEDURE DIVISION. DISPLAY 'ENV11'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'ENV11' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ENV11"
        DISPLAY "FAIL: want [ENV11] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

