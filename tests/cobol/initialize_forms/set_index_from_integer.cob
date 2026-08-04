*> vybe-test: cobol/initialize_forms/set_index_from_integer
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC X OCCURS 10 TIMES INDEXED BY IDX.
PROCEDURE DIVISION.
    SET IDX TO 1.
    STOP RUN.

