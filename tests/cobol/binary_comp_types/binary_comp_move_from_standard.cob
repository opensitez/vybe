*> vybe-test: cobol/binary_comp_types/binary_comp_move_from_standard
*> origin: languages/cobol/tests/cobol/test_binary_comp_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC 9(5) VALUE 12345.
01 DST PIC 9(5) COMP VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE SRC TO DST.
    DISPLAY DST.
    MOVE SPACES TO WS-VYBE-L
    STRING DST DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "12345"
        DISPLAY "FAIL: want [12345] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

