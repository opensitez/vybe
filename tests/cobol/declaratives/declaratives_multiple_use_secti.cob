*> vybe-test: cobol/declaratives/declaratives_multiple_use_sections
*> origin: languages/cobol/tests/cobol/test_declaratives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT file-a ASSIGN TO "a.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
           SELECT file-b ASSIGN TO "b.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD file-a.
       01 rec-a PIC X(80).
       FD file-b.
       01 rec-b PIC X(80).
       PROCEDURE DIVISION.
       DECLARATIVES.
       file-a-error SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON file-a.
           DISPLAY "file-a error".
       file-b-error SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON file-b.
           DISPLAY "file-b error".
       END DECLARATIVES.
       main-para SECTION.
           STOP RUN.

