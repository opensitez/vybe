*> vybe-test: cobol/initialize_forms/initialize_and_set_in_sequence
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(4) VALUE 9999.
01 FLAG PIC X VALUE "N".
    88 DONE VALUE "Y".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INITIALIZE N.
    SET DONE TO TRUE.
    IF DONE
        DISPLAY N
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING N DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0000"
        DISPLAY "FAIL: want [0000] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

