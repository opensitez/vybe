*> vybe-test: cobol/new_features/perform_thru_3
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    PERFORM STEP-A THRU STEP-C.
    STOP RUN.
STEP-A.
    DISPLAY "A".
STEP-B.
    DISPLAY "B".
STEP-C.
    DISPLAY "C".

