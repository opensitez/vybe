*> vybe-test: cobol/string_table_matrix/search_all_table_compiles
*> origin: languages/cobol/tests/cobol/test_string_table_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TAB.
   05 E OCCURS 4 TIMES ASCENDING KEY IS K INDEXED BY I.
      10 K PIC 9(2).
PROCEDURE DIVISION.
    MOVE 1 TO K(1).
    MOVE 2 TO K(2).
    MOVE 3 TO K(3).
    MOVE 4 TO K(4).
    SEARCH ALL E WHEN K(I) = 3 DISPLAY "Y" END-SEARCH.
    STOP RUN.

