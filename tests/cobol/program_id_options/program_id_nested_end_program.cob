*> vybe-test: cobol/program_id_options/program_id_nested_end_program
*> origin: languages/cobol/tests/cobol/test_program_id_options.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. outer.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-x PIC 9 VALUE 1.
       PROCEDURE DIVISION.
           DISPLAY ws-x
           STOP RUN.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. inner.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-y PIC 9 VALUE 2.
       PROCEDURE DIVISION.
           DISPLAY ws-y
           GOBACK.
       END PROGRAM inner.

       END PROGRAM outer.

