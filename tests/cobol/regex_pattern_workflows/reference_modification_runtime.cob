*> vybe-test: cobol/regex_pattern_workflows/reference_modification_runtime
*> origin: languages/cobol/tests/cobol/test_regex_pattern_workflows.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-T PIC X(12) VALUE "HELLOWORLD".
01 WS-S PIC X(5).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE WS-T(6:5) TO WS-S.
    DISPLAY WS-S.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "WORLD"
        DISPLAY "FAIL: want [WORLD] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

