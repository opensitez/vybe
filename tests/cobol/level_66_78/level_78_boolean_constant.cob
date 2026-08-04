*> vybe-test: cobol/level_66_78/level_78_boolean_constant
*> origin: languages/cobol/tests/cobol/test_level_66_78.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 TRUE-VAL       VALUE "Y".
       78 FALSE-VAL      VALUE "N".
       01 ws-active      PIC X VALUE FALSE-VAL.
       PROCEDURE DIVISION.
           MOVE TRUE-VAL TO ws-active
           IF ws-active = TRUE-VAL
               DISPLAY "Active"
           END-IF
           STOP RUN.

