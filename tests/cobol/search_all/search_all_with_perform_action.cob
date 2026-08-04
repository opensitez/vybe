*> vybe-test: cobol/search_all/search_all_with_perform_action
*> origin: languages/cobol/tests/cobol/test_search_all.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-product OCCURS 8 TIMES
               ASCENDING KEY IS ws-prod-id
               INDEXED BY ws-idx.
               10 ws-prod-id   PIC 9(5).
               10 ws-prod-name PIC X(20).
               10 ws-price     PIC 9(5)V99.
       01 ws-found-price PIC 9(5)V99 VALUE 0.
       01 ws-found-flag  PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE 10001 TO ws-prod-id(1) MOVE  5.99 TO ws-price(1)
           MOVE 10002 TO ws-prod-id(2) MOVE 12.50 TO ws-price(2)
           MOVE 10003 TO ws-prod-id(3) MOVE  3.25 TO ws-price(3)
           MOVE 10004 TO ws-prod-id(4) MOVE 99.00 TO ws-price(4)
           MOVE 10005 TO ws-prod-id(5) MOVE 14.75 TO ws-price(5)
           MOVE 10006 TO ws-prod-id(6) MOVE  7.49 TO ws-price(6)
           MOVE 10007 TO ws-prod-id(7) MOVE 22.00 TO ws-price(7)
           MOVE 10008 TO ws-prod-id(8) MOVE  1.99 TO ws-price(8)
           SEARCH ALL ws-product
               AT END
                   MOVE "N" TO ws-found-flag
               WHEN ws-prod-id(ws-idx) = 10005
                   MOVE ws-price(ws-idx) TO ws-found-price
                   MOVE "Y" TO ws-found-flag
           END-SEARCH
           IF ws-found-flag = "Y"
               DISPLAY ws-found-price
           ELSE
               DISPLAY "not found"
           END-IF
           STOP RUN.

