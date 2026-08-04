*> vybe-test: cobol/search_all/search_all_descending_key
*> origin: languages/cobol/tests/cobol/test_search_all.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 5 TIMES
               DESCENDING KEY IS ws-priority
               INDEXED BY ws-idx.
               10 ws-priority PIC 9.
               10 ws-task     PIC X(15).
       01 ws-found PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE 9 TO ws-priority(1)  MOVE "Critical" TO ws-task(1)
           MOVE 7 TO ws-priority(2)  MOVE "High"     TO ws-task(2)
           MOVE 5 TO ws-priority(3)  MOVE "Medium"   TO ws-task(3)
           MOVE 3 TO ws-priority(4)  MOVE "Low"      TO ws-task(4)
           MOVE 1 TO ws-priority(5)  MOVE "Deferred" TO ws-task(5)
           SEARCH ALL ws-entry
               AT END MOVE "N" TO ws-found
               WHEN ws-priority(ws-idx) = 5
                   MOVE "Y" TO ws-found
           END-SEARCH
           DISPLAY ws-found
           STOP RUN.

