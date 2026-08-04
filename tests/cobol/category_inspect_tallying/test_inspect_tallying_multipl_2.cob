*> vybe-test: cobol/category_inspect_tallying/test_inspect_tallying_multiple_conditions
*> origin: languages/cobol/tests/cobol/test_category_inspect_tallying.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 S1 PIC X(5) VALUE 'ABXAB'. 01 T1 PIC 9 VALUE 0. PROCEDURE DIVISION. INSPECT S1 TALLYING T1 FOR ALL 'A' BEFORE INITIAL 'X' ALL 'B' AFTER INITIAL 'X'. DISPLAY T1.
    MOVE SPACES TO WS-VYBE-L
    STRING T1 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2"
        DISPLAY "FAIL: want [2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

