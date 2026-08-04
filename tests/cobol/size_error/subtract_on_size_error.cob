*> vybe-test: cobol/size_error/subtract_on_size_error
*> origin: languages/cobol/tests/cobol/test_size_error.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-unsigned PIC 9 VALUE 0.
       01 ws-err      PIC X VALUE "N".
       PROCEDURE DIVISION.
           SUBTRACT 5 FROM ws-unsigned
               ON SIZE ERROR MOVE "Y" TO ws-err
           END-SUBTRACT
           DISPLAY ws-err
           STOP RUN.

