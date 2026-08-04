*> vybe-test: cobol/linage/linage_lines_at_bottom_only
*> origin: languages/cobol/tests/cobol/test_linage.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT rpt-file ASSIGN TO "rpt.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD rpt-file
           LINAGE IS 60 LINES
           LINES AT BOTTOM 4.
       01 rpt-line PIC X(132).
       PROCEDURE DIVISION.
           OPEN OUTPUT rpt-file
           MOVE "Detail line" TO rpt-line
           WRITE rpt-line
           CLOSE rpt-file
           STOP RUN.

