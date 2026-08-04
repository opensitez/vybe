*> vybe-test: cobol/conditions_extended/perform_until_condition_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I PIC 9(2) VALUE 0.
PROCEDURE DIVISION.
    PERFORM UNTIL WS-I >= 3
        ADD 1 TO WS-I
    END-PERFORM.
    STOP RUN.

