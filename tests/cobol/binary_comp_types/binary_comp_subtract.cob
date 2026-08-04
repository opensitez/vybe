*> vybe-test: cobol/binary_comp_types/binary_comp_subtract
*> origin: languages/cobol/tests/cobol/test_binary_comp_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(8) COMP VALUE 500.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    SUBTRACT 250 FROM N.
    DISPLAY N.
    MOVE SPACES TO WS-VYBE-L
    STRING N DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "250"
        DISPLAY "FAIL: want [250] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

