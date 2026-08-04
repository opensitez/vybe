*> vybe-test: cobol/category_inspect_replacing/test_insp_repl_all_before_initial
*> origin: languages/cobol/tests/cobol/test_category_inspect_replacing.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 S PIC X(5) VALUE 'A*AAA'. PROCEDURE DIVISION. INSPECT S REPLACING ALL 'A' BY 'X' BEFORE INITIAL '*'. DISPLAY S.
    MOVE SPACES TO WS-VYBE-L
    STRING S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "X*AAA"
        DISPLAY "FAIL: want [X*AAA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

