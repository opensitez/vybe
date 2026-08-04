*> vybe-test: cobol/final_features/add_corresponding
*> origin: languages/cobol/tests/cobol/test_final_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC.
   05 AMT PIC 9(5) VALUE 100.
01 DST.
   05 AMT PIC 9(5) VALUE 50.
PROCEDURE DIVISION.
    ADD CORRESPONDING SRC TO DST.
    STOP RUN.

