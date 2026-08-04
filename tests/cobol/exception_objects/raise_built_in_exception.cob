*> vybe-test: cobol/exception_objects/raise_built_in_exception
*> origin: languages/cobol/tests/cobol/test_exception_objects.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-caught PIC X VALUE "N".
       PROCEDURE DIVISION.
           RAISE EXCEPTION EC-PROGRAM-ARG-OMITTED
           DISPLAY ws-caught
           STOP RUN.

