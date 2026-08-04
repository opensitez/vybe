*> vybe-test: cobol/category_condition_logic/test_cond_and_elseif_chain
*> origin: languages/cobol/tests/cobol/test_category_condition_logic.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 A PIC 9 VALUE 2. PROCEDURE DIVISION. IF A = 1 DISPLAY 'A' ELSE IF A = 2 DISPLAY 'B' ELSE DISPLAY 'C' END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING 'A' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "B"
        DISPLAY "FAIL: want [B] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. END-IF. STOP RUN.

