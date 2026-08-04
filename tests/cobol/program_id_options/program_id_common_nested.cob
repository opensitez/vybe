*> vybe-test: cobol/program_id_options/program_id_common_nested
*> origin: languages/cobol/tests/cobol/test_program_id_options.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. outer-prog.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-shared PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           CALL "common-util"
           DISPLAY ws-shared
           STOP RUN.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. common-util IS COMMON.
       DATA DIVISION.
       PROCEDURE DIVISION.
           DISPLAY "common utility called"
           GOBACK.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. inner-prog.
       PROCEDURE DIVISION.
           CALL "common-util"
           GOBACK.
       END PROGRAM inner-prog.

       END PROGRAM common-util.
       END PROGRAM outer-prog.

