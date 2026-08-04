*> vybe-test: cobol/linage/linage_counter_access
*> origin: languages/cobol/tests/cobol/test_linage.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT cnt-file ASSIGN TO "cnt.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD cnt-file
           LINAGE IS 25 LINES
           WITH FOOTING AT 23.
       01 cnt-line PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-line-no PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           OPEN OUTPUT cnt-file
           PERFORM 10 TIMES
               MOVE LINAGE-COUNTER TO ws-line-no
               STRING "Line " DELIMITED SIZE
                      ws-line-no DELIMITED SIZE
                      INTO cnt-line
               WRITE cnt-line
           END-PERFORM
           CLOSE cnt-file
           STOP RUN.

