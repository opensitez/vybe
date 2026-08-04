*> vybe-test: cobol/linage/linage_with_footing
*> origin: languages/cobol/tests/cobol/test_linage.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT rpt ASSIGN TO "rpt.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD rpt
           LINAGE IS 56 LINES
           WITH FOOTING AT 54.
       01 rpt-line PIC X(132).
       PROCEDURE DIVISION.
           OPEN OUTPUT rpt
           MOVE "Report line 1" TO rpt-line
           WRITE rpt-line
           MOVE "Report line 2" TO rpt-line
           WRITE rpt-line
           CLOSE rpt
           STOP RUN.

