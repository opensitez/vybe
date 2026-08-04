*> vybe-test: cobol/size_error/multiply_giving_on_size_error
*> origin: languages/cobol/tests/cobol/test_size_error.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-a      PIC 99  VALUE 50.
       01 ws-b      PIC 99  VALUE 50.
       01 ws-result PIC 999 VALUE 0.
       01 ws-err    PIC X   VALUE "N".
       PROCEDURE DIVISION.
           MULTIPLY ws-a BY ws-b GIVING ws-result
               ON SIZE ERROR MOVE "Y" TO ws-err
           END-MULTIPLY
           DISPLAY ws-err
           STOP RUN.

