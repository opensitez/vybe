*> vybe-test: cobol/initialize_forms/initialize_table_element
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC 9(3) OCCURS 5 TIMES.
PROCEDURE DIVISION.
    INITIALIZE T.
    STOP RUN.

