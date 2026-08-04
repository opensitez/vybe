*> vybe-test: cobol/category_condition_logic/test_cond_nested_if_implicit_scope
*> origin: languages/cobol/tests/cobol/test_category_condition_logic.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 A PIC 9 VALUE 1. 01 B PIC 9 VALUE 2. PROCEDURE DIVISION. IF A = 1 IF B = 2 DISPLAY 'Y' END-IF END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING 'Y' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Y"
        DISPLAY "FAIL: want [Y] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

