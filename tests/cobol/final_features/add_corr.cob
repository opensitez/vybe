*> vybe-test: cobol/final_features/add_corr
*> origin: languages/cobol/tests/cobol/test_final_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A.
   05 X PIC 9(5) VALUE 10.
01 B.
   05 X PIC 9(5) VALUE 20.
PROCEDURE DIVISION.
    ADD CORR A TO B.
    STOP RUN.

