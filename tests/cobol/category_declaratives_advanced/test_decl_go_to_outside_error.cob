*> vybe-test: cobol/category_declaratives_advanced/test_decl_go_to_outside_error
*> origin: languages/cobol/tests/cobol/test_category_declaratives_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DECLARATIVES. D SECTION. USE AFTER ERROR ON INPUT. D-PARA. GO TO M-PARA. END DECLARATIVES. M SECTION. M-PARA. DISPLAY 'OK'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'OK' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

