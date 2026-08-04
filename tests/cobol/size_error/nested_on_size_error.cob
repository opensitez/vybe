*> vybe-test: cobol/size_error/nested_on_size_error
*> origin: languages/cobol/tests/cobol/test_size_error.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-x   PIC 9  VALUE 9.
       01 ws-y   PIC 9  VALUE 9.
       01 ws-err PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           ADD 5 TO ws-x
               ON SIZE ERROR
                   ADD 1 TO ws-err
                   ADD 5 TO ws-y
                       ON SIZE ERROR ADD 1 TO ws-err
                   END-ADD
           END-ADD
           DISPLAY ws-err
           STOP RUN.

