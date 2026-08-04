*> vybe-test: cobol/category_declaratives_use_before_reporting/test_use_before_reporting_multiple
*> origin: languages/cobol/tests/cobol/test_category_declaratives_use_before_reporting.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. DATA DIVISION. REPORT SECTION. RD R1. 01 G1 TYPE IS DETAIL. 01 G2 TYPE IS DETAIL. PROCEDURE DIVISION. DECLARATIVES. S1 SECTION. USE BEFORE REPORTING G1 G2. S1-PARA. DISPLAY 'BEF'.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING 'BEF' DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "MAIN"
                DISPLAY "FAIL at 1 want [MAIN] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE. END DECLARATIVES. MAIN SECTION. DISPLAY 'MAIN'.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING 'MAIN' DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "MAIN"
                DISPLAY "FAIL at 1 want [MAIN] got [" WS-VYBE-L "]"
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

