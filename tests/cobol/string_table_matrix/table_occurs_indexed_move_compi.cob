*> vybe-test: cobol/string_table_matrix/table_occurs_indexed_move_compiles
*> origin: languages/cobol/tests/cobol/test_string_table_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TBL.
   05 E OCCURS 5 TIMES INDEXED BY I.
      10 K PIC 9(2).
PROCEDURE DIVISION.
    SET I TO 1.
    MOVE 12 TO K(I).
    STOP RUN.

