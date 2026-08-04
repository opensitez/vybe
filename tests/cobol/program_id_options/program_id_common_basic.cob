*> vybe-test: cobol/program_id_options/program_id_common_basic
*> origin: languages/cobol/tests/cobol/test_program_id_options.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. host-prog.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-count PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           ADD 1 TO ws-count
           DISPLAY ws-count
           STOP RUN.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. shared-sub IS COMMON.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-msg PIC X(20) VALUE "shared utility".
       PROCEDURE DIVISION.
           DISPLAY ws-msg
           GOBACK.

       END PROGRAM shared-sub.
       END PROGRAM host-prog.

