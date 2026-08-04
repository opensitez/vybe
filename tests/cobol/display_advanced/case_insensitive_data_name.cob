*> vybe-test: cobol/display_advanced/case_insensitive_data_name
*> origin: languages/cobol/tests/cobol/test_display_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNTER PIC 9(3) VALUE 42.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY ws-counter.
    MOVE SPACES TO WS-VYBE-L
    STRING ws-counter DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "42"
        DISPLAY "FAIL: want [42] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    Move 7 To Ws-Counter.
    DISPLAY WS-COUNTER.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-COUNTER DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "7"
        DISPLAY "FAIL: want [7] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

