*> vybe-test: cobol/usage_clause/test_usage_native_binary
*> origin: languages/cobol/tests/cobol/test_usage_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-BIN PIC 9(4) USAGE IS COMP-5.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE 0 TO WS-BIN.
    ADD 1 TO WS-BIN.
    DISPLAY WS-BIN.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-BIN DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1"
        DISPLAY "FAIL: want [1] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

