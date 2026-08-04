*> vybe-test: cobol/printing_and_io/display_literal_runtime_prints_exact_text
*> origin: languages/cobol/tests/cobol/test_printing_and_io.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY "COBOL-IO".
    MOVE SPACES TO WS-VYBE-L
    STRING "COBOL-IO" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "COBOL-IO"
        DISPLAY "FAIL: want [COBOL-IO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

