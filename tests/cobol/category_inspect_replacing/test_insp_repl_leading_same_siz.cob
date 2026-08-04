*> vybe-test: cobol/category_inspect_replacing/test_insp_repl_leading_same_size
*> origin: languages/cobol/tests/cobol/test_category_inspect_replacing.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 S PIC X(5) VALUE 'AABAA'. PROCEDURE DIVISION. INSPECT S REPLACING LEADING 'A' BY 'X'. DISPLAY S.
    MOVE SPACES TO WS-VYBE-L
    STRING S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "XXBAA"
        DISPLAY "FAIL: want [XXBAA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

