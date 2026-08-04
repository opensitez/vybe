*> vybe-test: cobol/final_features/perform_test_before
*> origin: languages/cobol/tests/cobol/test_final_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    PERFORM WITH TEST BEFORE UNTIL WS-I >= 5
        ADD 1 TO WS-I
    END-PERFORM.
    STOP RUN.

