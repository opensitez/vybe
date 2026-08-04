*> vybe-test: cobol/string_and_unstring_extended/inspect_tallying_for_leading_zeroes_counts_prefix
*> origin: languages/cobol/tests/cobol/test_string_and_unstring_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(8) VALUE "0001234".
01 WS-CNT PIC 9(3) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    INSPECT WS-TXT TALLYING WS-CNT FOR LEADING "0".
    DISPLAY WS-CNT.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-CNT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "003"
        DISPLAY "FAIL: want [003] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

