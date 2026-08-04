*> vybe-test: cobol/figurative_constants_extended/figurative_constant_low_value_is_comparable
*> origin: languages/cobol/tests/cobol/test_figurative_constants_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-KEY PIC X(4) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE LOW-VALUES TO WS-KEY.
    IF WS-KEY = LOW-VALUES
        DISPLAY "LOW"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "LOW" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "LOW"
        DISPLAY "FAIL: want [LOW] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

