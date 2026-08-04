*> vybe-test: cobol/category_intrinsic_function_when_compiled/test_when_compiled_basic
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_when_compiled.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X(21). PROCEDURE DIVISION. MOVE FUNCTION WHEN-COMPILED TO V. IF V(1:4) > '2020' DISPLAY 'Y' END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING 'Y' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Y"
        DISPLAY "FAIL: want [Y] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

