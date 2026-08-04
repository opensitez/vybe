*> vybe-test: cobol/strings_and_text/inspect_replacing_rewrites_requested_characters
*> origin: languages/cobol/tests/cobol/test_strings_and_text.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(10) VALUE "ABCA".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INSPECT WS-TXT REPLACING FIRST "A" BY "Z".
    DISPLAY WS-TXT.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-TXT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ZBCA"
        DISPLAY "FAIL: want [ZBCA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

