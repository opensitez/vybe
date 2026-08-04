*> vybe-test: cobol/program_id_options/program_id_initial_reset_behavior
*> origin: languages/cobol/tests/cobol/test_program_id_options.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. reset-test INITIAL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val PIC 99 VALUE 10.
       01 ws-txt PIC X(10) VALUE "original".
       PROCEDURE DIVISION.
           ADD 5 TO ws-val
           MOVE "changed" TO ws-txt
           DISPLAY ws-val
           DISPLAY ws-txt
           STOP RUN.

