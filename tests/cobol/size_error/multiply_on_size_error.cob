*> vybe-test: cobol/size_error/multiply_on_size_error
*> origin: languages/cobol/tests/cobol/test_size_error.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9 VALUE 9.
       01 ws-err    PIC X VALUE "N".
       PROCEDURE DIVISION.
           MULTIPLY 9 BY ws-result
               ON SIZE ERROR MOVE "Y" TO ws-err
           END-MULTIPLY
           DISPLAY ws-err
           STOP RUN.

