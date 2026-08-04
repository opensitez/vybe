*> vybe-test: cobol/repository/repository_class_and_interface
*> origin: languages/cobol/tests/cobol/test_repository.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS     Shape     AS "Shape"
           INTERFACE Drawable  AS "Drawable".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-shape OBJECT REFERENCE Shape.
       PROCEDURE DIVISION.
           DISPLAY "class and interface in repository"
           STOP RUN.

