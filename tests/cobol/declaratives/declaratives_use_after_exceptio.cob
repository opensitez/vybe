*> vybe-test: cobol/declaratives/declaratives_use_after_exception_input
*> origin: languages/cobol/tests/cobol/test_declaratives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT in-file ASSIGN TO "in.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD in-file.
       01 in-rec PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-err PIC X VALUE "N".
       PROCEDURE DIVISION.
       DECLARATIVES.
       in-err SECTION.
           USE AFTER STANDARD EXCEPTION PROCEDURE ON INPUT.
           MOVE "Y" TO ws-err.
       END DECLARATIVES.
       main-para SECTION.
           STOP RUN.

