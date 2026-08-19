*> vybe-test: cobol/perform_and_evaluate_extended/perform_varying_with_step_and_limit
*> origin: languages/cobol/tests/cobol/test_perform_and_evaluate_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I PIC 9 VALUE 0.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
PROCEDURE DIVISION.

    PERFORM VARYING WS-I FROM 1 BY 2 UNTIL WS-I > 5
        DISPLAY WS-I
    END-PERFORM.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING WS-I DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "7"
                DISPLAY "FAIL at 1 want [7] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "3"
                DISPLAY "FAIL at 2 want [3] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "5"
                DISPLAY "FAIL at 3 want [5] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 3 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.

    IF WS-VYBE-I NOT = 3
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 3"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

