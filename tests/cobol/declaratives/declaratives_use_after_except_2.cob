*> vybe-test: cobol/declaratives/declaratives_use_after_exception_output
*> origin: languages/cobol/tests/cobol/test_declaratives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT out-file ASSIGN TO "out.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD out-file.
       01 out-rec PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-err PIC X VALUE "N".
       PROCEDURE DIVISION.
       DECLARATIVES.
       out-err SECTION.
           USE AFTER STANDARD EXCEPTION PROCEDURE ON OUTPUT.
           MOVE "Y" TO ws-err.
       END DECLARATIVES.
       main-para SECTION.
           STOP RUN.

