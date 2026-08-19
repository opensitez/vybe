*> vybe-test: cobol/string_functions_intrinsic/intrinsic_length_numeric_field
*> origin: languages/cobol/tests/cobol/test_string_functions_intrinsic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(8) VALUE 0.
01 L PIC 9(4) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    COMPUTE L = FUNCTION LENGTH(N).
    DISPLAY L.
    MOVE SPACES TO WS-VYBE-L
    STRING L DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0008"
        DISPLAY "FAIL: want [0008] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

