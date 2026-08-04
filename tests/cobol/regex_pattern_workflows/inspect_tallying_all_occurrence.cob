*> vybe-test: cobol/regex_pattern_workflows/inspect_tallying_all_occurrences_runtime
*> origin: languages/cobol/tests/cobol/test_regex_pattern_workflows.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(20) VALUE "ABCAABC".
01 WS-CNT PIC 9(2) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INSPECT WS-TEXT TALLYING WS-CNT FOR ALL "A".
    DISPLAY WS-CNT.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-CNT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "3"
        DISPLAY "FAIL: want [3] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

