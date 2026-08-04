*> vybe-test: cobol/size_error/add_on_size_error
*> origin: languages/cobol/tests/cobol/test_size_error.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-small PIC 9 VALUE 9.
       01 ws-overflow PIC X VALUE "N".
       PROCEDURE DIVISION.
           ADD 1 TO ws-small
               ON SIZE ERROR MOVE "Y" TO ws-overflow
           END-ADD
           DISPLAY ws-overflow
           STOP RUN.

