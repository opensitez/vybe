*> vybe-test: cobol/repository/repository_mixed
*> origin: languages/cobol/tests/cobol/test_repository.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           FUNCTION ALL INTRINSIC
           CLASS Connection AS "Connection"
           CLASS ResultSet  AS "ResultSet"
           INTERFACE Closeable AS "Closeable".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-conn OBJECT REFERENCE Connection.
       01 ws-rs   OBJECT REFERENCE ResultSet.
       01 ws-len  PIC 99.
       PROCEDURE DIVISION.
           COMPUTE ws-len = LENGTH("hello")
           DISPLAY ws-len
           STOP RUN.

