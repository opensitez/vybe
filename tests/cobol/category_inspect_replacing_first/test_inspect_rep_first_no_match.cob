*> vybe-test: cobol/category_inspect_replacing_first/test_inspect_rep_first_no_match
*> origin: languages/cobol/tests/cobol/test_category_inspect_replacing_first.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 S1 PIC X(5) VALUE 'BBBBB'. PROCEDURE DIVISION. INSPECT S1 REPLACING FIRST 'A' BY 'X'. DISPLAY S1.
    MOVE SPACES TO WS-VYBE-L
    STRING S1 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BBBBB"
        DISPLAY "FAIL: want [BBBBB] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

