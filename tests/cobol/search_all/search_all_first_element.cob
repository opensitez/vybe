*> vybe-test: cobol/search_all/search_all_first_element
*> origin: languages/cobol/tests/cobol/test_search_all.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 10 TIMES
               ASCENDING KEY IS ws-key
               INDEXED BY ws-idx.
               10 ws-key  PIC 9(2).
               10 ws-data PIC X(5).
       01 ws-found PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE 10 TO ws-key(1)   MOVE 20 TO ws-key(2)
           MOVE 30 TO ws-key(3)   MOVE 40 TO ws-key(4)
           MOVE 50 TO ws-key(5)   MOVE 60 TO ws-key(6)
           MOVE 70 TO ws-key(7)   MOVE 80 TO ws-key(8)
           MOVE 90 TO ws-key(9)   MOVE 99 TO ws-key(10)
           SEARCH ALL ws-entry
               AT END MOVE "N" TO ws-found
               WHEN ws-key(ws-idx) = 10
                   MOVE "Y" TO ws-found
           END-SEARCH
           DISPLAY ws-found
           STOP RUN.

