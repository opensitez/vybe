*> vybe-test: cobol/initialize_forms/set_index_to_length_of_table
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC X OCCURS 5 TIMES INDEXED BY IX.
PROCEDURE DIVISION.
    SET IX TO 5.
    STOP RUN.

