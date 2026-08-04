*> vybe-test: cobol/search_all/search_all_nested_table
*> origin: languages/cobol/tests/cobol/test_search_all.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-outer.
           05 ws-region OCCURS 3 TIMES.
               10 ws-region-code PIC X(3).
               10 ws-city-table.
                   15 ws-city OCCURS 4 TIMES
                       ASCENDING KEY IS ws-city-id
                       INDEXED BY ws-ci.
                       20 ws-city-id   PIC 9(3).
                       20 ws-city-name PIC X(15).
       01 ws-found PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE "NE" TO ws-region-code(1)
           MOVE 101 TO ws-city-id(1, 1)
           MOVE 102 TO ws-city-id(1, 2)
           MOVE 103 TO ws-city-id(1, 3)
           MOVE 104 TO ws-city-id(1, 4)
           SEARCH ALL ws-city(1)
               AT END MOVE "N" TO ws-found
               WHEN ws-city-id(1, ws-ci) = 103
                   MOVE "Y" TO ws-found
           END-SEARCH
           DISPLAY ws-found
           STOP RUN.

