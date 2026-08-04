*> vybe-test: cobol/category_declaratives_advanced/test_decl_nested_program_global
*> origin: languages/cobol/tests/cobol/test_category_declaratives_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DECLARATIVES. D SECTION. USE GLOBAL AFTER ERROR ON INPUT. D-PARA. DISPLAY 'ERR'.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING 'ERR' DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "SUB OK"
                DISPLAY "FAIL at 1 want [SUB OK] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE. END DECLARATIVES. M SECTION. CALL 'SUB'. STOP RUN. IDENTIFICATION DIVISION. PROGRAM-ID. SUB. PROCEDURE DIVISION. S SECTION. DISPLAY 'SUB OK'.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING 'SUB OK' DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "SUB OK"
                DISPLAY "FAIL at 1 want [SUB OK] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE. EXIT PROGRAM. END PROGRAM SUB.
    IF WS-VYBE-I NOT = 1
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 1"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

