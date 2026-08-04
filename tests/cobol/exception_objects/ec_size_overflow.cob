*> vybe-test: cobol/exception_objects/ec_size_overflow
*> origin: languages/cobol/tests/cobol/test_exception_objects.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-n PIC 9 VALUE 9.
       PROCEDURE DIVISION.
           ADD 5 TO ws-n
               ON SIZE ERROR RAISE EXCEPTION EC-SIZE-OVERFLOW
           END-ADD
           STOP RUN.

