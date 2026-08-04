*> vybe-test: cobol/size_error/compute_on_size_error
*> origin: languages/cobol/tests/cobol/test_size_error.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9 VALUE 0.
       01 ws-err    PIC X VALUE "N".
       PROCEDURE DIVISION.
           COMPUTE ws-result = 999 * 999
               ON SIZE ERROR MOVE "Y" TO ws-err
           END-COMPUTE
           DISPLAY ws-err
           STOP RUN.

