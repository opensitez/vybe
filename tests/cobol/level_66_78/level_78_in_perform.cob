*> vybe-test: cobol/level_66_78/level_78_in_perform
*> origin: languages/cobol/tests/cobol/test_level_66_78.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 LOOP-COUNT     VALUE 5.
       01 ws-idx         PIC 9.
       01 ws-sum         PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM VARYING ws-idx FROM 1 BY 1
               UNTIL ws-idx > LOOP-COUNT
               ADD ws-idx TO ws-sum
           END-PERFORM
           DISPLAY ws-sum
           STOP RUN.

