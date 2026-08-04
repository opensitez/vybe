*> vybe-test: cobol/search_all/search_all_basic
*> origin: languages/cobol/tests/cobol/test_search_all.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 5 TIMES
               ASCENDING KEY IS ws-code
               INDEXED BY ws-idx.
               10 ws-code  PIC 9(3).
               10 ws-label PIC X(10).
       01 ws-found PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE 100 TO ws-code(1)  MOVE "Alpha"    TO ws-label(1)
           MOVE 200 TO ws-code(2)  MOVE "Beta"     TO ws-label(2)
           MOVE 300 TO ws-code(3)  MOVE "Gamma"    TO ws-label(3)
           MOVE 400 TO ws-code(4)  MOVE "Delta"    TO ws-label(4)
           MOVE 500 TO ws-code(5)  MOVE "Epsilon"  TO ws-label(5)
           SEARCH ALL ws-entry
               AT END MOVE "N" TO ws-found
               WHEN ws-code(ws-idx) = 300
                   MOVE "Y" TO ws-found
           END-SEARCH
           DISPLAY ws-found
           STOP RUN.

