*> vybe-test: cobol/nested_if_else/if_compute_in_outer_then_nested_check
*> origin: languages/cobol/tests/cobol/test_nested_if_else.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 6.
01 B PIC 9(3) VALUE 7.
01 PROD PIC 9(5) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    COMPUTE PROD = A * B.
    IF PROD > 40
        IF PROD < 50
            DISPLAY "RANGE"
        ELSE
            DISPLAY "HIGH"
        END-IF
    ELSE
        DISPLAY "LOW"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "RANGE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "RANGE"
        DISPLAY "FAIL: want [RANGE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

