*> vybe-test: cobol/conversion_and_coercion_extended/move_numeric_to_alpha_runtime
*> origin: languages/cobol/tests/cobol/test_conversion_and_coercion_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SRC PIC 9(3) VALUE 42.
01 WS-DST PIC X(5) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE WS-SRC TO WS-DST.
    DISPLAY WS-DST.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-DST DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "042"
        DISPLAY "FAIL: want [042] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

