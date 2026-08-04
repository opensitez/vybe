*> vybe-test: cobol/category_inspect_tallying/test_inspect_tallying_all_before
*> origin: languages/cobol/tests/cobol/test_category_inspect_tallying.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 S1 PIC X(5) VALUE 'AABAA'. 01 T PIC 9 VALUE 0. PROCEDURE DIVISION. INSPECT S1 TALLYING T FOR ALL 'A' BEFORE INITIAL 'B'. DISPLAY T.
    MOVE SPACES TO WS-VYBE-L
    STRING T DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2"
        DISPLAY "FAIL: want [2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

