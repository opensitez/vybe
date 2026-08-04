*> vybe-test: cobol/table_search_binary/search_linear_condition_greater_than
*> origin: languages/cobol/tests/cobol/test_table_search_binary.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 N PIC 9(3) OCCURS 5 TIMES INDEXED BY IX.
PROCEDURE DIVISION.
    MOVE 10 TO N(1). MOVE 20 TO N(2). MOVE 30 TO N(3).
    MOVE 40 TO N(4). MOVE 50 TO N(5).
    SET IX TO 1.
    SEARCH N
        AT END DISPLAY "NOT FOUND"
        WHEN N(IX) > 25 DISPLAY "FIRST OVER 25"
    END-SEARCH.
    STOP RUN.

