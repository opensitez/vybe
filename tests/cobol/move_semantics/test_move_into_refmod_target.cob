*> vybe-test: cobol/move_semantics/test_move_into_refmod_target
*> origin: languages/cobol/tests/cobol/test_move_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(6) VALUE "AABBCC".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE "XX" TO WS-TEXT(3:2).
    DISPLAY WS-TEXT.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-TEXT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AAXXCC"
        DISPLAY "FAIL: want [AAXXCC] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

