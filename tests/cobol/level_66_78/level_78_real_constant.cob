*> vybe-test: cobol/level_66_78/level_78_real_constant
*> origin: languages/cobol/tests/cobol/test_level_66_78.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 PI             VALUE 3.14159265.
       78 E              VALUE 2.71828182.
       01 ws-result      PIC 9(3)V9(8).
       PROCEDURE DIVISION.
           MOVE PI TO ws-result
           DISPLAY ws-result
           STOP RUN.

