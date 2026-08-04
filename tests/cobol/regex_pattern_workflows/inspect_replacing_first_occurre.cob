*> vybe-test: cobol/regex_pattern_workflows/inspect_replacing_first_occurrence_runtime
*> origin: languages/cobol/tests/cobol/test_regex_pattern_workflows.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(10) VALUE "BANANA".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INSPECT WS-TEXT REPLACING FIRST "A" BY "X".
    DISPLAY WS-TEXT.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-TEXT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BXNANA"
        DISPLAY "FAIL: want [BXNANA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

