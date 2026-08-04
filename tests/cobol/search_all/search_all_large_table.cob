*> vybe-test: cobol/search_all/search_all_large_table
*> origin: languages/cobol/tests/cobol/test_search_all.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 100 TIMES
               ASCENDING KEY IS ws-code
               INDEXED BY ws-idx.
               10 ws-code  PIC 9(5).
               10 ws-value PIC X(10).
       01 ws-found PIC X VALUE "N".
       01 ws-i PIC 9(3).
       PROCEDURE DIVISION.
           PERFORM VARYING ws-i FROM 1 BY 1 UNTIL ws-i > 100
               MULTIPLY ws-i BY 10 GIVING ws-code(ws-i)
           END-PERFORM
           SEARCH ALL ws-entry
               AT END MOVE "N" TO ws-found
               WHEN ws-code(ws-idx) = 500
                   MOVE "Y" TO ws-found
           END-SEARCH
           DISPLAY ws-found
           STOP RUN.

