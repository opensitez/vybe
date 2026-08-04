*> vybe-test: cobol/binary_numeric_types/comp_values_display_result
*> origin: languages/cobol/tests/cobol/test_binary_numeric_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. BN11.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 A PIC S9(4) COMP VALUE 10.
01 B PIC S9(4) COMP VALUE 5.
01 C PIC S9(4) COMP.
PROCEDURE DIVISION.
    ADD A TO B
    MOVE B TO C
    DISPLAY C
    STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING C DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "15"
        DISPLAY "FAIL: want [15] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

