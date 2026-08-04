*> vybe-test: cobol/category_inspect_replacing/test_insp_converting
*> origin: languages/cobol/tests/cobol/test_category_inspect_replacing.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 S PIC X(5) VALUE 'ABCDE'. PROCEDURE DIVISION. INSPECT S CONVERTING 'ACE' TO '123'. DISPLAY S.
    MOVE SPACES TO WS-VYBE-L
    STRING S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1B2D3"
        DISPLAY "FAIL: want [1B2D3] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

