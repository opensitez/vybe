*> vybe-test: cobol/initialize_forms/set_condition_inside_if
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 1.
01 FLAG PIC X VALUE "N".
    88 FOUND VALUE "Y".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF N > 0
        SET FOUND TO TRUE
    END-IF.
    IF FOUND
        DISPLAY "YES"
    ELSE
        DISPLAY "NO"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "YES" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "YES"
        DISPLAY "FAIL: want [YES] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

