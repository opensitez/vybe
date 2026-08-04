*> vybe-test: cobol/repository/repository_class_basic
*> origin: languages/cobol/tests/cobol/test_repository.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS MyClass AS "MyClass".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-obj OBJECT REFERENCE MyClass.
       PROCEDURE DIVISION.
           DISPLAY "repository class declared"
           STOP RUN.

