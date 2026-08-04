*> vybe-test: cobol/string_and_unstring_extended/inspect_replacing_first_character_changes_only_first_match
*> origin: languages/cobol/tests/cobol/test_string_and_unstring_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(6) VALUE "AAAAAA".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    INSPECT WS-TXT REPLACING FIRST "A" BY "B".
    DISPLAY WS-TXT.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-TXT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BAAAAA"
        DISPLAY "FAIL: want [BAAAAA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

