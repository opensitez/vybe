*> vybe-test: cobol/final_features/set_88_true
*> origin: languages/cobol/tests/cobol/test_final_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATUS PIC X(1).
   88 IS-ACTIVE VALUE "A".
   88 IS-INACTIVE VALUE "I".
PROCEDURE DIVISION.
    SET IS-ACTIVE TO TRUE.
    SET IS-INACTIVE TO FALSE.
    STOP RUN.

