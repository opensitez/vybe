*> vybe-test: cobol/program_id_options/program_id_initial_basic
*> origin: languages/cobol/tests/cobol/test_program_id_options.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test INITIAL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-counter PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           ADD 1 TO ws-counter
           DISPLAY ws-counter
           STOP RUN.

