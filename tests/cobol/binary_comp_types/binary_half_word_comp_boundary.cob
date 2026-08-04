*> vybe-test: cobol/binary_comp_types/binary_half_word_comp_boundary
*> origin: languages/cobol/tests/cobol/test_binary_comp_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(4) COMP VALUE 9999.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    ADD 1 TO N.
    DISPLAY N.
    MOVE SPACES TO WS-VYBE-L
    STRING N DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "10000"
        DISPLAY "FAIL: want [10000] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

