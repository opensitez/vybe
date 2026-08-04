*> vybe-test: cobol/level_66_78/level_78_used_in_if
*> origin: languages/cobol/tests/cobol/test_level_66_78.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 PASS-MARK      VALUE 50.
       78 DISTINCTION    VALUE 85.
       01 ws-score       PIC 999 VALUE 92.
       PROCEDURE DIVISION.
           IF ws-score >= DISTINCTION
               DISPLAY "Distinction"
           ELSE IF ws-score >= PASS-MARK
               DISPLAY "Pass"
           ELSE
               DISPLAY "Fail"
           END-IF
           STOP RUN.

