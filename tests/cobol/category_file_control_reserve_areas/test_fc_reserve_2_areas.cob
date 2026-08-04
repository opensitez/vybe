*> vybe-test: cobol/category_file_control_reserve_areas/test_fc_reserve_2_areas
*> origin: languages/cobol/tests/cobol/test_category_file_control_reserve_areas.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. INPUT-OUTPUT SECTION. FILE-CONTROL. SELECT F ASSIGN TO 'a' RESERVE 2 AREAS. PROCEDURE DIVISION. DISPLAY 'OK'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'OK' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

