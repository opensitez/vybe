*> vybe-test: cobol/data_division_expanded/redefines_share_buffer_when_nesting_changes
*> origin: languages/cobol/tests/cobol/test_data_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-BUF PIC X(2).
01 WS-GRP REDEFINES WS-BUF.
   05 WS-CH1 PIC X.
   05 WS-CH2 PIC X.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "AB" TO WS-GRP.
    MOVE WS-GRP TO WS-BUF.
    DISPLAY WS-BUF.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-BUF DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AB"
        DISPLAY "FAIL: want [AB] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

