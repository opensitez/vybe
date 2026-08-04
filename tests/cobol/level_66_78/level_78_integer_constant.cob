*> vybe-test: cobol/level_66_78/level_78_integer_constant
*> origin: languages/cobol/tests/cobol/test_level_66_78.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 MAX-SIZE       VALUE 100.
       78 MIN-SIZE       VALUE 1.
       01 ws-count       PIC 999.
       PROCEDURE DIVISION.
           MOVE MAX-SIZE TO ws-count
           DISPLAY ws-count
           STOP RUN.

