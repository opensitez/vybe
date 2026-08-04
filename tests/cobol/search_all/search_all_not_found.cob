*> vybe-test: cobol/search_all/search_all_not_found
*> origin: languages/cobol/tests/cobol/test_search_all.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 5 TIMES
               ASCENDING KEY IS ws-id
               INDEXED BY ws-idx.
               10 ws-id  PIC 9(4).
               10 ws-val PIC X(5).
       01 ws-result PIC X(10) VALUE "not found".
       PROCEDURE DIVISION.
           MOVE 1010 TO ws-id(1)
           MOVE 2020 TO ws-id(2)
           MOVE 3030 TO ws-id(3)
           MOVE 4040 TO ws-id(4)
           MOVE 5050 TO ws-id(5)
           SEARCH ALL ws-entry
               AT END MOVE "missing"   TO ws-result
               WHEN ws-id(ws-idx) = 9999
                   MOVE "found"        TO ws-result
           END-SEARCH
           DISPLAY ws-result
           STOP RUN.

