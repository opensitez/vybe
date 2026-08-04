*> vybe-test: cobol/program_id_options/program_id_with_end_program
*> origin: languages/cobol/tests/cobol/test_program_id_options.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. bounded-prog.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val PIC 99 VALUE 42.
       PROCEDURE DIVISION.
           DISPLAY ws-val
           STOP RUN.
       END PROGRAM bounded-prog.

