*> vybe-test: cobol/conversion_and_coercion_extended/compute_numeric_then_move_to_alphanumeric_runtime
*> origin: languages/cobol/tests/cobol/test_conversion_and_coercion_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(2) VALUE 8.
01 B PIC 9(2) VALUE 5.
01 N PIC 9(3) VALUE 0.
01 T PIC X(5).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    COMPUTE N = A + B.
    MOVE N TO T.
    DISPLAY T.
    MOVE SPACES TO WS-VYBE-L
    STRING T DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "13"
        DISPLAY "FAIL: want [13] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

