*> vybe-test: cobol/figurative_constants_extended/figurative_constant_high_values_fill_field
*> origin: languages/cobol/tests/cobol/test_figurative_constants_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-BUF PIC X(4) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE HIGH-VALUES TO WS-BUF.
    DISPLAY WS-BUF.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-BUF DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ɕɕɕɕ"
        DISPLAY "FAIL: want [ɕɕɕɕ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

