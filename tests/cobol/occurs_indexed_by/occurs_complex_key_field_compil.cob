*> vybe-test: cobol/occurs_indexed_by/occurs_complex_key_field_compiles
*> origin: languages/cobol/tests/cobol/test_occurs_indexed_by.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 LOOKUP.
   05 ENTRY OCCURS 20 TIMES ASCENDING KEY CODE-VAL INDEXED BY LK-IDX.
      10 CODE-VAL PIC X(4).
      10 DESC-VAL PIC X(20).
PROCEDURE DIVISION.
    SET LK-IDX TO 1.
    MOVE "AAAA" TO CODE-VAL(LK-IDX).
    STOP RUN.

