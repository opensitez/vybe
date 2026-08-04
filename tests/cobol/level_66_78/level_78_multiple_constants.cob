*> vybe-test: cobol/level_66_78/level_78_multiple_constants
*> origin: languages/cobol/tests/cobol/test_level_66_78.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 STATUS-OK      VALUE 0.
       78 STATUS-WARN    VALUE 1.
       78 STATUS-ERROR   VALUE 2.
       78 STATUS-FATAL   VALUE 3.
       01 ws-status      PIC 9.
       PROCEDURE DIVISION.
           MOVE STATUS-OK TO ws-status
           EVALUATE ws-status
               WHEN STATUS-OK    DISPLAY "OK"
               WHEN STATUS-WARN  DISPLAY "Warning"
               WHEN STATUS-ERROR DISPLAY "Error"
               WHEN STATUS-FATAL DISPLAY "Fatal"
           END-EVALUATE
           STOP RUN.

