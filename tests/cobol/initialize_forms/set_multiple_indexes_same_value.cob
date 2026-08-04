*> vybe-test: cobol/initialize_forms/set_multiple_indexes_same_value
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T1.
   05 E1 PIC X OCCURS 5 TIMES INDEXED BY IX1.
01 T2.
   05 E2 PIC X OCCURS 5 TIMES INDEXED BY IX2.
PROCEDURE DIVISION.
    SET IX1 IX2 TO 1.
    STOP RUN.

