*> vybe-test: cobol/category_condition_logic/test_cond_complex_and_or_parens
*> origin: languages/cobol/tests/cobol/test_category_condition_logic.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 A PIC 9 VALUE 0. 01 B PIC 9 VALUE 1. 01 C PIC 9 VALUE 1. PROCEDURE DIVISION. IF A = 1 AND (B = 1 OR C = 1) DISPLAY 'Y' ELSE DISPLAY 'N' END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING 'Y' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "N"
        DISPLAY "FAIL: want [N] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

