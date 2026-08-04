*> vybe-test: cobol/collections_dictionaries/keyed_table_decl_compiles
*> origin: languages/cobol/tests/cobol/test_collections_dictionaries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 MAP.
   05 ENTRY OCCURS 5 TIMES ASCENDING KEY IS K INDEXED BY I.
      10 K PIC X(5).
      10 V PIC X(10).
PROCEDURE DIVISION.
    MOVE "A" TO K(1).
    STOP RUN.

