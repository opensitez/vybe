*> vybe-test: cobol/inspect_converting/inspect_converting_single_char_pair
*> origin: languages/cobol/tests/cobol/test_inspect_converting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "HELLO".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INSPECT S CONVERTING "L" TO "R".
    DISPLAY S.
    MOVE SPACES TO WS-VYBE-L
    STRING S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HERRO     "
        DISPLAY "FAIL: want [HERRO     ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

