*> vybe-test: cobol/search_all/search_all_varying_occurs
*> origin: languages/cobol/tests/cobol/test_search_all.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-count PIC 99 VALUE 5.
       01 ws-table.
           05 ws-entry OCCURS 1 TO 20 TIMES
               DEPENDING ON ws-count
               ASCENDING KEY IS ws-val
               INDEXED BY ws-idx.
               10 ws-val PIC 9(3).
       01 ws-found PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE 5 TO ws-count
           MOVE 10 TO ws-val(1)
           MOVE 20 TO ws-val(2)
           MOVE 30 TO ws-val(3)
           MOVE 40 TO ws-val(4)
           MOVE 50 TO ws-val(5)
           SEARCH ALL ws-entry
               AT END MOVE "N" TO ws-found
               WHEN ws-val(ws-idx) = 30
                   MOVE "Y" TO ws-found
           END-SEARCH
           DISPLAY ws-found
           STOP RUN.

