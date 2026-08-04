*> vybe-test: cobol/alter_stop/stop_literal_string
*> origin: languages/cobol/tests/cobol/test_alter_stop.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-flag PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE "Y" TO ws-flag
           DISPLAY ws-flag
           STOP "Press Enter to continue"
           STOP RUN.

