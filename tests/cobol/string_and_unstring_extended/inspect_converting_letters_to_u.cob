*> vybe-test: cobol/string_and_unstring_extended/inspect_converting_letters_to_upper_case
*> origin: languages/cobol/tests/cobol/test_string_and_unstring_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(8) VALUE "abc123".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    INSPECT WS-TXT CONVERTING "abc" TO "ABC".
    DISPLAY WS-TXT.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-TXT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABC123  "
        DISPLAY "FAIL: want [ABC123  ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

