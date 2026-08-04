*> vybe-test: cobol/linage/linage_full_specification
*> origin: languages/cobol/tests/cobol/test_linage.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT full-rpt ASSIGN TO "full.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD full-rpt
           LINAGE IS 60 LINES
           WITH FOOTING AT 58
           LINES AT TOP 3
           LINES AT BOTTOM 3.
       01 full-line PIC X(132).
       PROCEDURE DIVISION.
           OPEN OUTPUT full-rpt
           MOVE "Report header line" TO full-line
           WRITE full-line AFTER ADVANCING PAGE
           MOVE "Detail line 1" TO full-line
           WRITE full-line AFTER ADVANCING 1 LINE
           MOVE "Detail line 2" TO full-line
           WRITE full-line AFTER ADVANCING 1 LINE
           CLOSE full-rpt
           STOP RUN.

