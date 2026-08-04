*> vybe-test: cobol/math_builtins/compute_mixed_add_sub_runtime_expected_result
*> origin: languages/cobol/tests/cobol/test_math_builtins.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 99 VALUE 20.
01 B PIC 99 VALUE 5.
01 C PIC 99 VALUE 3.
01 R PIC 99 VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    COMPUTE R = A - B + C.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "18"
        DISPLAY "FAIL: want [18] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

