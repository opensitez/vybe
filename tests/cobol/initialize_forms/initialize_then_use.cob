*> vybe-test: cobol/initialize_forms/initialize_then_use
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(4) VALUE 9999.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INITIALIZE N.
    ADD 42 TO N.
    DISPLAY N.
    MOVE SPACES TO WS-VYBE-L
    STRING N DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0042"
        DISPLAY "FAIL: want [0042] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

