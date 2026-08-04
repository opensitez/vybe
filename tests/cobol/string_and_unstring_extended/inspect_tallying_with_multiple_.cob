*> vybe-test: cobol/string_and_unstring_extended/inspect_tallying_with_multiple_patterns
*> origin: languages/cobol/tests/cobol/test_string_and_unstring_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(12) VALUE "ABBAABBA".
01 WS-CNT PIC 9(3) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    INSPECT WS-TXT TALLYING WS-CNT FOR ALL "A".
    DISPLAY WS-CNT.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-CNT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "004"
        DISPLAY "FAIL: want [004] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

