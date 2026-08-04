*> vybe-test: cobol/category_declaratives_advanced/test_decl_priority
*> origin: languages/cobol/tests/cobol/test_category_declaratives_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. INPUT-OUTPUT SECTION. FILE-CONTROL. SELECT F ASSIGN TO 'a'. DATA DIVISION. FILE SECTION. FD F. 01 R PIC X. PROCEDURE DIVISION. DECLARATIVES. D1 SECTION. USE AFTER ERROR ON F. D1-PARA. DISPLAY 'F-ERR'.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING 'F-ERR' DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "F-ERR"
                DISPLAY "FAIL at 1 want [F-ERR] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE. D2 SECTION. USE AFTER ERROR ON INPUT. D2-PARA. DISPLAY 'IN-ERR'.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING 'IN-ERR' DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "F-ERR"
                DISPLAY "FAIL at 1 want [F-ERR] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE. END DECLARATIVES. M SECTION. OPEN INPUT F. STOP RUN.
    IF WS-VYBE-I NOT = 1
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 1"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

