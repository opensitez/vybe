*> vybe-test: cobol/regex_pattern_workflows/pattern_preprocessing_with_intrinsics_runtime
*> origin: languages/cobol/tests/cobol/test_regex_pattern_workflows.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(20) VALUE "  AbC123  ".
01 WS-NORM PIC X(20).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE FUNCTION TRIM(WS-TEXT) TO WS-NORM.
    MOVE FUNCTION LOWER-CASE(WS-NORM) TO WS-NORM.
    DISPLAY WS-NORM.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-NORM DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "abc123"
        DISPLAY "FAIL: want [abc123] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

