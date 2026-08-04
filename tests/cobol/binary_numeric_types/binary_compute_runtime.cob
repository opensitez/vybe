*> vybe-test: cobol/binary_numeric_types/binary_compute_runtime
*> origin: languages/cobol/tests/cobol/test_binary_numeric_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. BN15.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 A PIC S9(4) COMP VALUE 25.
01 B PIC S9(4) COMP VALUE 5.
01 C PIC S9(4) COMP.
PROCEDURE DIVISION.
    COMPUTE C = A / B
    DISPLAY C
    STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING C DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "5"
        DISPLAY "FAIL: want [5] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

