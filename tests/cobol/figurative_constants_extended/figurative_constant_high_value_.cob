*> vybe-test: cobol/figurative_constants_extended/figurative_constant_high_value_is_comparable
*> origin: languages/cobol/tests/cobol/test_figurative_constants_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-KEY PIC X(4) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE HIGH-VALUES TO WS-KEY.
    IF WS-KEY = HIGH-VALUES
        DISPLAY "HIGH"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "HIGH" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HIGH"
        DISPLAY "FAIL: want [HIGH] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

