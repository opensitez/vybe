*> vybe-test: cobol/table_search_binary/search_all_performs_action_on_found_field
*> origin: languages/cobol/tests/cobol/test_table_search_binary.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 LOOKUP.
   05 L OCCURS 50 TIMES ASCENDING KEY L-KEY INDEXED BY LI.
      10 L-KEY PIC 9(5).
      10 L-VALUE PIC X(20).
01 FOUND-VAL PIC X(20) VALUE SPACES.
PROCEDURE DIVISION.
    SEARCH ALL L
        AT END
            MOVE "MISSING" TO FOUND-VAL
        WHEN L-KEY(LI) = 12345
            MOVE L-VALUE(LI) TO FOUND-VAL
    END-SEARCH.
    DISPLAY FOUND-VAL.
    STOP RUN.

