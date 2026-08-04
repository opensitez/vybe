*> vybe-test: cobol/program_id_options/program_id_initial_with_call
*> origin: languages/cobol/tests/cobol/test_program_id_options.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. main-prog.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           CALL "fresh-sub" USING ws-result
           DISPLAY ws-result
           STOP RUN.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. fresh-sub INITIAL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-internal PIC 99 VALUE 0.
       LINKAGE SECTION.
       01 ls-result PIC 99.
       PROCEDURE DIVISION USING ls-result.
           ADD 1 TO ws-internal
           MOVE ws-internal TO ls-result
           GOBACK.

