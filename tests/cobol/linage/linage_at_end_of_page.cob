*> vybe-test: cobol/linage/linage_at_end_of_page
*> origin: languages/cobol/tests/cobol/test_linage.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT eop-file ASSIGN TO "eop.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD eop-file
           LINAGE IS 10 LINES
           WITH FOOTING AT 9.
       01 eop-line PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-idx PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           OPEN OUTPUT eop-file
           PERFORM VARYING ws-idx FROM 1 BY 1 UNTIL ws-idx > 25
               MOVE ws-idx TO eop-line
               WRITE eop-line
                   AT END-OF-PAGE
                       MOVE "--- page break ---" TO eop-line
                       WRITE eop-line AFTER ADVANCING PAGE
               END-WRITE
           END-PERFORM
           CLOSE eop-file
           STOP RUN.

