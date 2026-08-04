*> vybe-test: cobol/figurative_constants_extended/figurative_constant_move_all_repeats_pattern
*> origin: languages/cobol/tests/cobol/test_figurative_constants_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-LINE PIC X(6) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE ALL "-" TO WS-LINE.
    DISPLAY WS-LINE.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-LINE DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "------"
        DISPLAY "FAIL: want [------] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

