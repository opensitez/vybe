*> vybe-test: cobol/math_builtins/compute_addition_runtime_displays_expected_sum
*> origin: languages/cobol/tests/cobol/test_math_builtins.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(2) VALUE 12.
01 WS-B PIC 9(2) VALUE 7.
01 WS-C PIC 9(3) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    COMPUTE WS-C = WS-A + WS-B.
    DISPLAY WS-C.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-C DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "019"
        DISPLAY "FAIL: want [019] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

