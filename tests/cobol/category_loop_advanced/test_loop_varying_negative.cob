*> vybe-test: cobol/category_loop_advanced/test_loop_varying_negative
*> origin: languages/cobol/tests/cobol/test_category_loop_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. LOOP-NEG.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
       01 I PIC S9 VALUE 0.
       01 TOTAL PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM VARYING I FROM 5 BY -1 UNTIL I < 1
              ADD 1 TO TOTAL
           END-PERFORM.
           DISPLAY TOTAL.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING TOTAL DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "05"
                DISPLAY "FAIL at 1 want [05] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
           STOP RUN.
    
    IF WS-VYBE-I NOT = 1
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 1"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

