*> vybe-test: cobol/perform_advanced/test_perform_exit_perform
*> origin: languages/cobol/tests/cobol/test_perform_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I PIC 9.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
PROCEDURE DIVISION.

    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 5
        IF WS-I = 3
            EXIT PERFORM
        END-IF
        DISPLAY WS-I
    END-PERFORM.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING WS-I DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "3"
                DISPLAY "FAIL at 1 want [3] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "2"
                DISPLAY "FAIL at 2 want [2] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 2 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.

    IF WS-VYBE-I NOT = 2
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 2"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

