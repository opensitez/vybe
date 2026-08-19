*> vybe-test: cobol/inspect_converting/inspect_tallying_leading_zeros
*> origin: languages/cobol/tests/cobol/test_inspect_converting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(8) VALUE "000123".
01 C PIC 9(2) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INSPECT S TALLYING C FOR LEADING "0".
    DISPLAY C.
    MOVE SPACES TO WS-VYBE-L
    STRING C DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "03"
        DISPLAY "FAIL: want [03] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

