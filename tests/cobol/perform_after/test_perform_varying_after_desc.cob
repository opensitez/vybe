*> vybe-test: cobol/perform_after/test_perform_varying_after_descending
*> origin: languages/cobol/tests/cobol/test_perform_after.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I PIC 9.
01 WS-J PIC 9.
01 WS-TOTAL PIC 99 VALUE 0.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
PROCEDURE DIVISION.

    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 2
      AFTER WS-J FROM 3 BY -1 UNTIL WS-J < 1
        ADD 1 TO WS-TOTAL
    END-PERFORM
    DISPLAY WS-TOTAL.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING WS-TOTAL DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "06"
                DISPLAY "FAIL at 1 want [06] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.

    IF WS-VYBE-I NOT = 1
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 1"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

