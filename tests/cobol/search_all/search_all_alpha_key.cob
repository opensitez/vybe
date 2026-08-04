*> vybe-test: cobol/search_all/search_all_alpha_key
*> origin: languages/cobol/tests/cobol/test_search_all.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 5 TIMES
               ASCENDING KEY IS ws-name
               INDEXED BY ws-idx.
               10 ws-name  PIC X(10).
               10 ws-score PIC 99.
       01 ws-result PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           MOVE "Alice"    TO ws-name(1)  MOVE 85 TO ws-score(1)
           MOVE "Bob"      TO ws-name(2)  MOVE 92 TO ws-score(2)
           MOVE "Charlie"  TO ws-name(3)  MOVE 78 TO ws-score(3)
           MOVE "Diana"    TO ws-name(4)  MOVE 95 TO ws-score(4)
           MOVE "Eve"      TO ws-name(5)  MOVE 88 TO ws-score(5)
           SEARCH ALL ws-entry
               AT END MOVE 0 TO ws-result
               WHEN ws-name(ws-idx) = "Charlie"
                   MOVE ws-score(ws-idx) TO ws-result
           END-SEARCH
           DISPLAY ws-result
           STOP RUN.

