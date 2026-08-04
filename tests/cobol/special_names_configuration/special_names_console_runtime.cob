*> vybe-test: cobol/special_names_configuration/special_names_console_runtime
*> origin: languages/cobol/tests/cobol/test_special_names_configuration.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CONSOLE IS CRT.
PROCEDURE DIVISION.
    DISPLAY "CRT".
    MOVE SPACES TO WS-VYBE-L
    STRING "CRT" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "CRT"
        DISPLAY "FAIL: want [CRT] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

