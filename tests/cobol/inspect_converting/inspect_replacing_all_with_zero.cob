*> vybe-test: cobol/inspect_converting/inspect_replacing_all_with_zero
*> origin: languages/cobol/tests/cobol/test_inspect_converting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(6) VALUE "X1X2X3".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INSPECT S REPLACING ALL "X" BY "0".
    DISPLAY S.
    MOVE SPACES TO WS-VYBE-L
    STRING S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "010203"
        DISPLAY "FAIL: want [010203] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

