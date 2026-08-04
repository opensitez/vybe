*> vybe-test: cobol/level_66_78/level_78_mixed_with_66
*> origin: languages/cobol/tests/cobol/test_level_66_78.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 RECORD-SIZE    VALUE 80.
       01 full-record.
           05 rec-header PIC X(10).
           05 rec-body   PIC X(60).
           05 rec-footer PIC X(10).
       66 rec-content RENAMES rec-header THRU rec-body.
       01 ws-len PIC 999.
       PROCEDURE DIVISION.
           MOVE RECORD-SIZE TO ws-len
           MOVE "HDR" TO rec-header
           DISPLAY ws-len
           STOP RUN.

