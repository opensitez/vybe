*> vybe-test: cobol/figurative_constants_extended/figurative_constant_move_all_repeats_multiple_characters
*> origin: languages/cobol/tests/cobol/test_figurative_constants_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(8) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE ALL "12" TO WS-TXT.
    DISPLAY WS-TXT.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-TXT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "12121212"
        DISPLAY "FAIL: want [12121212] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

