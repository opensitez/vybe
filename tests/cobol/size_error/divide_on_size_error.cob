*> vybe-test: cobol/size_error/divide_on_size_error
*> origin: languages/cobol/tests/cobol/test_size_error.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-dividend PIC 9 VALUE 5.
       01 ws-divisor  PIC 9 VALUE 0.
       01 ws-err      PIC X VALUE "N".
       PROCEDURE DIVISION.
           DIVIDE ws-divisor INTO ws-dividend
               ON SIZE ERROR MOVE "Y" TO ws-err
           END-DIVIDE
           DISPLAY ws-err
           STOP RUN.

