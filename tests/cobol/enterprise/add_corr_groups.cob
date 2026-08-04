*> vybe-test: cobol/enterprise/add_corr_groups
*> origin: languages/cobol/tests/cobol/test_enterprise.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC.
   05 AMT1 PIC 9(5) VALUE 100.
   05 AMT2 PIC 9(5) VALUE 200.
01 DST.
   05 AMT1 PIC 9(5) VALUE 0.
   05 AMT2 PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    ADD CORRESPONDING SRC TO DST.
    STOP RUN.

