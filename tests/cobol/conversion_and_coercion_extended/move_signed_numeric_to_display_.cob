*> vybe-test: cobol/conversion_and_coercion_extended/move_signed_numeric_to_display_runtime
*> origin: languages/cobol/tests/cobol/test_conversion_and_coercion_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-S PIC S9(3) VALUE -12.
01 WS-D PIC X(6).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE WS-S TO WS-D.
    DISPLAY WS-D.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-D DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "012"
        DISPLAY "FAIL: want [012] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

