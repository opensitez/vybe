*> vybe-test: cobol/move_semantics_extended/move_all_to_alphanumeric_field_repeats_pattern
*> origin: languages/cobol/tests/cobol/test_move_semantics_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(6) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE ALL "AB" TO WS-TXT.
    DISPLAY WS-TXT.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-TXT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABABAB"
        DISPLAY "FAIL: want [ABABAB] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

