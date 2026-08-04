*> vybe-test: cobol/declaratives/declaratives_use_after_error_on_file
*> origin: languages/cobol/tests/cobol/test_declaratives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT my-file ASSIGN TO "test.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD my-file.
       01 my-rec PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-error-msg PIC X(50) VALUE SPACES.
       PROCEDURE DIVISION.
       DECLARATIVES.
       file-error SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON my-file.
           MOVE "File error occurred" TO ws-error-msg
           DISPLAY ws-error-msg.
       END DECLARATIVES.
       main-logic SECTION.
           OPEN INPUT my-file
           CLOSE my-file
           STOP RUN.

