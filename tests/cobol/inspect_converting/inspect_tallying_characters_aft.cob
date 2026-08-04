*> vybe-test: cobol/inspect_converting/inspect_tallying_characters_after_replacing
*> origin: languages/cobol/tests/cobol/test_inspect_converting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(6) VALUE "AABBCC".
01 C PIC 9(2) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INSPECT S REPLACING ALL "A" BY "X".
    INSPECT S TALLYING C FOR ALL "X".
    DISPLAY C.
    MOVE SPACES TO WS-VYBE-L
    STRING C DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2"
        DISPLAY "FAIL: want [2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

