*> vybe-test: cobol/occurs_indexed_by/occurs_set_from_index_to_integer
*> origin: languages/cobol/tests/cobol/test_occurs_indexed_by.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC X OCCURS 5 TIMES INDEXED BY IX.
01 POS PIC 9 VALUE 0.
PROCEDURE DIVISION.
    SET IX TO 3.
    SET POS TO IX.
    STOP RUN.

