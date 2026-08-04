*> vybe-test: cobol/inspect_converting/inspect_replacing_leading_zeros_runtime
*> origin: languages/cobol/tests/cobol/test_inspect_converting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(8) VALUE "00000042".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INSPECT S REPLACING LEADING "0" BY " ".
    DISPLAY S.
    MOVE SPACES TO WS-VYBE-L
    STRING S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "      42"
        DISPLAY "FAIL: want [      42] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

