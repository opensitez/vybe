*> vybe-test: cobol/initialize_forms/set_index_down_by
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC X OCCURS 10 TIMES INDEXED BY IDX.
PROCEDURE DIVISION.
    SET IDX TO 5.
    SET IDX DOWN BY 2.
    STOP RUN.

