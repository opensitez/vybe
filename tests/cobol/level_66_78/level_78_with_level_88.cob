*> vybe-test: cobol/level_66_78/level_78_with_level_88
*> origin: languages/cobol/tests/cobol/test_level_66_78.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 MAX-RETRIES    VALUE 3.
       01 ws-retry-count PIC 9 VALUE 0.
           88 max-reached VALUE MAX-RETRIES.
       PROCEDURE DIVISION.
           MOVE MAX-RETRIES TO ws-retry-count
           IF max-reached
               DISPLAY "Max retries reached"
           END-IF
           STOP RUN.

