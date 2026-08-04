*> vybe-test: cobol/table_search_binary/search_all_multiple_when_condition_compiles
*> origin: languages/cobol/tests/cobol/test_table_search_binary.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SORTED.
   05 ITEM OCCURS 20 TIMES ASCENDING KEY ITEM-KEY INDEXED BY SI.
      10 ITEM-KEY PIC 9(4).
      10 ITEM-VAL PIC X(10).
PROCEDURE DIVISION.
    SEARCH ALL ITEM
        AT END
            DISPLAY "END"
        WHEN ITEM-KEY(SI) = 42
            DISPLAY ITEM-VAL(SI)
        WHEN ITEM-KEY(SI) = 99
            DISPLAY "NINETY-NINE"
    END-SEARCH.
    STOP RUN.

