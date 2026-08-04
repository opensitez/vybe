*> vybe-test: cobol/figurative_constants/high_values_end_of_table_sentinel
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 5 TIMES.
               10 ws-code PIC X(5).
       01 ws-idx PIC 9.
       PROCEDURE DIVISION.
           MOVE HIGH-VALUES TO ws-code(1)
           MOVE HIGH-VALUES TO ws-code(2)
           MOVE HIGH-VALUES TO ws-code(3)
           MOVE HIGH-VALUES TO ws-code(4)
           MOVE HIGH-VALUES TO ws-code(5)
           PERFORM VARYING ws-idx FROM 1 BY 1
               UNTIL ws-idx > 5 OR ws-code(ws-idx) = HIGH-VALUES
               DISPLAY ws-code(ws-idx)
           END-PERFORM
           STOP RUN.

