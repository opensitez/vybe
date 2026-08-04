*> vybe-test: cobol/string_and_unstring_extended/inspect_replacing_with_characters
*> origin: languages/cobol/tests/cobol/test_string_and_unstring_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(8) VALUE "ABC123".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    INSPECT WS-TXT REPLACING CHARACTERS BY "X".
    DISPLAY WS-TXT.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-TXT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "XXXXXXXX"
        DISPLAY "FAIL: want [XXXXXXXX] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

