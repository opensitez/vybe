*> vybe-test: cobol/move_semantics/test_move_numeric_left_truncation
*> origin: languages/cobol/tests/cobol/test_move_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SRC PIC 9(5) VALUE 12345.
01 WS-DST PIC 9(3) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE WS-SRC TO WS-DST.
    DISPLAY WS-DST.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-DST DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "345"
        DISPLAY "FAIL: want [345] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

