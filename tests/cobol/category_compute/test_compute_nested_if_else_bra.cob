*> vybe-test: cobol/category_compute/test_compute_nested_if_else_branch
*> origin: languages/cobol/tests/cobol/test_category_compute.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 A PIC S9(4) VALUE 0. PROCEDURE DIVISION. IF 1 = 1 COMPUTE A = (10 + 20) * 2 - 5 END-IF DISPLAY A.
    MOVE SPACES TO WS-VYBE-L
    STRING A DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0055"
        DISPLAY "FAIL: want [0055] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

