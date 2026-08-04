*> vybe-test: cobol/conditions_extended/if_nested_condition_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 5.
PROCEDURE DIVISION.
    IF WS-A > 0
        IF WS-A < 10
            DISPLAY "A"
        END-IF
    END-IF.
    STOP RUN.

