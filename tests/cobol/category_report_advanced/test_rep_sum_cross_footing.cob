*> vybe-test: cobol/category_report_advanced/test_rep_sum_cross_footing
*> origin: languages/cobol/tests/cobol/test_category_report_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. REPORT SECTION. RD R. 01 TYPE IS CONTROL FOOTING FINAL. 05 S1 COL 1 PIC 9. 05 S2 COL 2 PIC 9. 05 S3 COL 3 PIC 9 SUM S1 S2. PROCEDURE DIVISION. DISPLAY 'OK'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'OK' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

