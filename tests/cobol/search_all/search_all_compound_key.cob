*> vybe-test: cobol/search_all/search_all_compound_key
*> origin: languages/cobol/tests/cobol/test_search_all.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 6 TIMES
               ASCENDING KEY IS ws-dept ws-emp-id
               INDEXED BY ws-idx.
               10 ws-dept   PIC X(3).
               10 ws-emp-id PIC 9(4).
               10 ws-salary PIC 9(6).
       01 ws-found PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE "ACC" TO ws-dept(1)  MOVE 1001 TO ws-emp-id(1)
           MOVE "ACC" TO ws-dept(2)  MOVE 1002 TO ws-emp-id(2)
           MOVE "ENG" TO ws-dept(3)  MOVE 2001 TO ws-emp-id(3)
           MOVE "ENG" TO ws-dept(4)  MOVE 2002 TO ws-emp-id(4)
           MOVE "MKT" TO ws-dept(5)  MOVE 3001 TO ws-emp-id(5)
           MOVE "MKT" TO ws-dept(6)  MOVE 3002 TO ws-emp-id(6)
           SEARCH ALL ws-entry
               AT END MOVE "N" TO ws-found
               WHEN ws-dept(ws-idx) = "ENG" AND
                    ws-emp-id(ws-idx) = 2002
                   MOVE "Y" TO ws-found
           END-SEARCH
           DISPLAY ws-found
           STOP RUN.

