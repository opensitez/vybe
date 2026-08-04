*> vybe-test: cobol/category_pointers_advanced/test_ptr_group_move
*> origin: languages/cobol/tests/cobol/test_category_pointers_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 G1. 05 P1 USAGE POINTER. 01 G2. 05 P2 USAGE POINTER. PROCEDURE DIVISION. SET P1 TO NULL. MOVE G1 TO G2. IF P2 = NULL DISPLAY 'Y' END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING 'Y' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Y"
        DISPLAY "FAIL: want [Y] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

