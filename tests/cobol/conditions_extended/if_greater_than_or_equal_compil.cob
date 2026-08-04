*> vybe-test: cobol/conditions_extended/if_greater_than_or_equal_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 10.
01 WS-B PIC 9(3) VALUE 10.
PROCEDURE DIVISION.
    IF WS-A >= WS-B
        DISPLAY "A"
    END-IF.
    STOP RUN.

