*> vybe-test: cobol/linage/linage_write_advancing_page
*> origin: languages/cobol/tests/cobol/test_linage.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT pg-file ASSIGN TO "pg.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD pg-file
           LINAGE IS 30 LINES
           WITH FOOTING AT 28.
       01 pg-line PIC X(80).
       PROCEDURE DIVISION.
           OPEN OUTPUT pg-file
           MOVE "Page 1 Header" TO pg-line
           WRITE pg-line AFTER ADVANCING PAGE
           MOVE "Page 1 body"   TO pg-line
           WRITE pg-line AFTER ADVANCING 1 LINE
           MOVE "Page 2 Header" TO pg-line
           WRITE pg-line AFTER ADVANCING PAGE
           CLOSE pg-file
           STOP RUN.

