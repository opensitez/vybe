*> vybe-test: cobol/category_declaratives_use_before_reporting/test_use_before_reporting_parse_10
*> origin: languages/cobol/tests/cobol/test_category_declaratives_use_before_reporting.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. DATA DIVISION. REPORT SECTION. RD R9. 01 G TYPE IS DETAIL. 05 COLUMN 1 PIC X VALUE 'I'. PROCEDURE DIVISION. DECLARATIVES. D9 SECTION. USE BEFORE REPORTING R9. D9-PARA. DISPLAY "OPEN".
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING "OPEN" DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "OK"
                DISPLAY "FAIL at 1 want [OK] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE. END DECLARATIVES. MAIN SECTION. DISPLAY 'OK'.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING 'OK' DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "OK"
                DISPLAY "FAIL at 1 want [OK] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE. STOP RUN.
    IF WS-VYBE-I NOT = 1
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 1"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

