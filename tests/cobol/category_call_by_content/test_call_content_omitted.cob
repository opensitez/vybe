*> vybe-test: cobol/category_call_by_content/test_call_content_omitted
*> origin: languages/cobol/tests/cobol/test_category_call_by_content.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. CALL 'S1' BY CONTENT OMITTED. STOP RUN. IDENTIFICATION DIVISION. PROGRAM-ID. S1. PROCEDURE DIVISION. DISPLAY 'OK'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'OK' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. EXIT PROGRAM. END PROGRAM S1.

