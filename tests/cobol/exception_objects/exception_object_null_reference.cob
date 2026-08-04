*> vybe-test: cobol/exception_objects/exception_object_null_reference
*> origin: languages/cobol/tests/cobol/test_exception_objects.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS EXCEPTION-OBJECT AS "EXCEPTION-OBJECT".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-obj OBJECT REFERENCE EXCEPTION-OBJECT VALUE NULL.
       PROCEDURE DIVISION.
           IF ws-obj = NULL
               DISPLAY "null object reference"
           END-IF
           STOP RUN.

