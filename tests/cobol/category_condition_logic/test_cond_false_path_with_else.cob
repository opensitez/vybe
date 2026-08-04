*> vybe-test: cobol/category_condition_logic/test_cond_false_path_with_else
*> origin: languages/cobol/tests/cobol/test_category_condition_logic.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 A PIC S9 VALUE -1. PROCEDURE DIVISION. IF A IS POSITIVE DISPLAY 'P' ELSE IF A IS NEGATIVE DISPLAY 'N' ELSE DISPLAY 'Z' END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING 'P' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "N"
        DISPLAY "FAIL: want [N] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. END-IF. STOP RUN.

