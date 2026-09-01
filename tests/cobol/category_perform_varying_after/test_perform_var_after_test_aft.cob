*> vybe-test: cobol/category_perform_varying_after/test_perform_var_after_test_after
*> origin: languages/cobol/tests/cobol/test_category_perform_varying_after.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0. 01 I PIC 9. 01 J PIC 9. PROCEDURE DIVISION. PERFORM M-PARA WITH TEST AFTER VARYING I FROM 1 BY 1 UNTIL I > 2 AFTER J FROM 1 BY 1 UNTIL J > 2. DISPLAY 'OK'.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING 'OK' DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "11"
                DISPLAY "FAIL at 1 want [11] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "12"
                DISPLAY "FAIL at 2 want [12] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "13"
                DISPLAY "FAIL at 3 want [13] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 4
            IF WS-VYBE-L NOT = "21"
                DISPLAY "FAIL at 4 want [21] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 5
            IF WS-VYBE-L NOT = "22"
                DISPLAY "FAIL at 5 want [22] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 6
            IF WS-VYBE-L NOT = "23"
                DISPLAY "FAIL at 6 want [23] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 7
            IF WS-VYBE-L NOT = "31"
                DISPLAY "FAIL at 7 want [31] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 8
            IF WS-VYBE-L NOT = "32"
                DISPLAY "FAIL at 8 want [32] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 9
            IF WS-VYBE-L NOT = "33"
                DISPLAY "FAIL at 9 want [33] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 10
            IF WS-VYBE-L NOT = "OK"
                DISPLAY "FAIL at 10 want [OK] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 10 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.    IF WS-VYBE-I NOT = 10
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 10"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
 STOP RUN. M-PARA. DISPLAY I J.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING I DELIMITED SIZE J DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "11"
                DISPLAY "FAIL at 1 want [11] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "12"
                DISPLAY "FAIL at 2 want [12] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "13"
                DISPLAY "FAIL at 3 want [13] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 4
            IF WS-VYBE-L NOT = "21"
                DISPLAY "FAIL at 4 want [21] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 5
            IF WS-VYBE-L NOT = "22"
                DISPLAY "FAIL at 5 want [22] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 6
            IF WS-VYBE-L NOT = "23"
                DISPLAY "FAIL at 6 want [23] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 7
            IF WS-VYBE-L NOT = "31"
                DISPLAY "FAIL at 7 want [31] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 8
            IF WS-VYBE-L NOT = "32"
                DISPLAY "FAIL at 8 want [32] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 9
            IF WS-VYBE-L NOT = "33"
                DISPLAY "FAIL at 9 want [33] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 10
            IF WS-VYBE-L NOT = "OK"
                DISPLAY "FAIL at 10 want [OK] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 10 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.

