*> vybe-test: cobol/binary_numeric_types/comp_multiply_remainder_runtime
*> origin: languages/cobol/tests/cobol/test_binary_numeric_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. BN13.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 A PIC S9(4) COMP VALUE 3.
01 B PIC S9(4) COMP VALUE 4.
01 C PIC S9(4) COMP.
PROCEDURE DIVISION.
    MULTIPLY A BY B GIVING C
    DISPLAY C
    STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING C DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "12"
        DISPLAY "FAIL: want [12] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

