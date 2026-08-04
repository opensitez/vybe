*> vybe-test: cobol/display_formatting/display_group_item_shows_raw_bytes
*> origin: languages/cobol/tests/cobol/test_display_formatting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 GRP.
   05 GA PIC X(3) VALUE "ABC".
   05 GB PIC X(3) VALUE "DEF".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY GRP.
    MOVE SPACES TO WS-VYBE-L
    STRING GRP DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABCDEF"
        DISPLAY "FAIL: want [ABCDEF] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

