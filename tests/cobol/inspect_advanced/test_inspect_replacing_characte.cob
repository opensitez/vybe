*> vybe-test: cobol/inspect_advanced/test_inspect_replacing_characters
*> origin: languages/cobol/tests/cobol/test_inspect_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STR PIC X(6) VALUE "SECRET".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    INSPECT WS-STR REPLACING CHARACTERS BY "*".
    DISPLAY WS-STR.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-STR DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "******"
        DISPLAY "FAIL: want [******] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

