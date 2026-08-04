*> vybe-test: cobol/alter_stop/stop_literal_in_conditional
*> origin: languages/cobol/tests/cobol/test_alter_stop.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-debug PIC X VALUE "Y".
       PROCEDURE DIVISION.
           IF ws-debug = "Y"
               STOP "Debug checkpoint reached"
           END-IF
           DISPLAY "continuing"
           STOP RUN.

