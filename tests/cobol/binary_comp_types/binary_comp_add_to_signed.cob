*> vybe-test: cobol/binary_comp_types/binary_comp_add_to_signed
*> origin: languages/cobol/tests/cobol/test_binary_comp_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC S9(5) COMP VALUE -100.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    ADD 200 TO N.
    DISPLAY N.
    MOVE SPACES TO WS-VYBE-L
    STRING N DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "+100"
        DISPLAY "FAIL: want [+100] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

