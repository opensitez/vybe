*> vybe-test: cobol/category_inspect_replacing/test_insp_tallying_leading
*> origin: languages/cobol/tests/cobol/test_category_inspect_replacing.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 S PIC X(5) VALUE 'AABAA'. 01 C PIC 9 VALUE 0. PROCEDURE DIVISION. INSPECT S TALLYING C FOR LEADING 'A'. DISPLAY C.
    MOVE SPACES TO WS-VYBE-L
    STRING C DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2"
        DISPLAY "FAIL: want [2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

