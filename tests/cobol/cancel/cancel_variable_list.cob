*> vybe-test: cobol/cancel/cancel_variable_list
*> origin: languages/cobol/tests/cobol/test_cancel.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-mods.
           05 ws-mod-1 PIC X(20) VALUE "module-x".
           05 ws-mod-2 PIC X(20) VALUE "module-y".
       PROCEDURE DIVISION.
           CALL ws-mod-1
           CALL ws-mod-2
           CANCEL ws-mod-1
           CANCEL ws-mod-2
           DISPLAY "freed"
           STOP RUN.

