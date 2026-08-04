*> vybe-test: cobol/repository/repository_class_hierarchy
*> origin: languages/cobol/tests/cobol/test_repository.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS Vehicle  AS "Vehicle"
           CLASS Car      AS "Car"
           CLASS Truck    AS "Truck"
           CLASS Fleet    AS "Fleet".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-vehicle OBJECT REFERENCE Vehicle.
       01 ws-car     OBJECT REFERENCE Car.
       01 ws-truck   OBJECT REFERENCE Truck.
       PROCEDURE DIVISION.
           DISPLAY "fleet management system"
           STOP RUN.

