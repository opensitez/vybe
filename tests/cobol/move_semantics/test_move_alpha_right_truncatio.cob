*> vybe-test: cobol/move_semantics/test_move_alpha_right_truncation
*> origin: languages/cobol/tests/cobol/test_move_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SRC PIC X(6) VALUE "ABCDEF".
01 WS-DST PIC X(3) VALUE "XXX".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE WS-SRC TO WS-DST.
    DISPLAY WS-DST.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-DST DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABC"
        DISPLAY "FAIL: want [ABC] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

