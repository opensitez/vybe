*> vybe-test: cobol/standard_compliant_core/add_giving_produces_expected_result
*> origin: languages/cobol/tests/cobol/test_standard_compliant_core.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(2) VALUE 7.
01 WS-B PIC 9(2) VALUE 8.
01 WS-R PIC 9(3) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    ADD WS-A WS-B GIVING WS-R.
    DISPLAY WS-R.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "015"
        DISPLAY "FAIL: want [015] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

