*> vybe-test: cobol/binary_numeric_types/binary_move_to_displayable_field_runtime
*> origin: languages/cobol/tests/cobol/test_binary_numeric_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. BN17.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 A PIC S9(4) COMP VALUE -12.
01 B PIC S9(4).
PROCEDURE DIVISION.
    MOVE A TO B
    DISPLAY B
    STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING B DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "-12"
        DISPLAY "FAIL: want [-12] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

