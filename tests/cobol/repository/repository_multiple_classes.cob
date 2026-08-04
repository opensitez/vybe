*> vybe-test: cobol/repository/repository_multiple_classes
*> origin: languages/cobol/tests/cobol/test_repository.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS Animal   AS "Animal"
           CLASS Dog      AS "Dog"
           CLASS Cat      AS "Cat".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-animal OBJECT REFERENCE Animal.
       01 ws-dog    OBJECT REFERENCE Dog.
       PROCEDURE DIVISION.
           DISPLAY "multi-class repository"
           STOP RUN.

