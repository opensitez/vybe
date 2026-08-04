*> vybe-test: cobol/category_file_control_advanced/test_fc_same_record_area
*> origin: languages/cobol/tests/cobol/test_category_file_control_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. INPUT-OUTPUT SECTION. I-O-CONTROL. SAME RECORD AREA FOR F1 F2. PROCEDURE DIVISION. DISPLAY 'OK'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'OK' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

