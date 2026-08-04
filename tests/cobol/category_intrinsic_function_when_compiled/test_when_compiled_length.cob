*> vybe-test: cobol/category_intrinsic_function_when_compiled/test_when_compiled_length
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_when_compiled.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION LENGTH(FUNCTION WHEN-COMPILED).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE LENGTH(FUNCTION DELIMITED SIZE WHEN-COMPILED) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "21"
        DISPLAY "FAIL: want [21] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

