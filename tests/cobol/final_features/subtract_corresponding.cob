*> vybe-test: cobol/final_features/subtract_corresponding
*> origin: languages/cobol/tests/cobol/test_final_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC.
   05 AMT PIC 9(5) VALUE 30.
01 DST.
   05 AMT PIC 9(5) VALUE 100.
PROCEDURE DIVISION.
    SUBTRACT CORRESPONDING SRC FROM DST.
    STOP RUN.

