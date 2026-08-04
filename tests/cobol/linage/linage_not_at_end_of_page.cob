*> vybe-test: cobol/linage/linage_not_at_end_of_page
*> origin: languages/cobol/tests/cobol/test_linage.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT nep-file ASSIGN TO "nep.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD nep-file
           LINAGE IS 20 LINES
           WITH FOOTING AT 18.
       01 nep-line PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-line-count PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           OPEN OUTPUT nep-file
           MOVE "Test line" TO nep-line
           WRITE nep-line
               AT END-OF-PAGE
                   ADD 1 TO ws-line-count
               NOT AT END-OF-PAGE
                   DISPLAY "still on page"
           END-WRITE
           CLOSE nep-file
           STOP RUN.

