*> vybe-test: cobol/search_all/search_all_index_after_search
*> origin: languages/cobol/tests/cobol/test_search_all.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 5 TIMES
               ASCENDING KEY IS ws-num
               INDEXED BY ws-idx.
               10 ws-num PIC 9(3).
       01 ws-idx-val PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           MOVE 100 TO ws-num(1)
           MOVE 200 TO ws-num(2)
           MOVE 300 TO ws-num(3)
           MOVE 400 TO ws-num(4)
           MOVE 500 TO ws-num(5)
           SEARCH ALL ws-entry
               AT END MOVE 0 TO ws-idx-val
               WHEN ws-num(ws-idx) = 300
                   SET ws-idx-val TO ws-idx
           END-SEARCH
           DISPLAY ws-idx-val
           STOP RUN.

