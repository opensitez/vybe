*> vybe-test: cobol/perform_out_of_line/perform_paragraph_until_uses_updated_ws
*> origin: languages/cobol/tests/cobol/test_perform_out_of_line.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TOTAL PIC 9(4) VALUE 0.
01 STEP  PIC 9(3) VALUE 100.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    PERFORM ADD-STEP UNTIL TOTAL >= 500.
    DISPLAY TOTAL.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING TOTAL DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "0500"
                DISPLAY "FAIL at 1 want [0500] got [" WS-VYBE-L "]"
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
ADD-STEP.
    ADD STEP TO TOTAL.
    STOP RUN.

