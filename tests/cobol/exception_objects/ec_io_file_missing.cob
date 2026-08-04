*> vybe-test: cobol/exception_objects/ec_io_file_missing
*> origin: languages/cobol/tests/cobol/test_exception_objects.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT data-file ASSIGN TO "nonexistent.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD data-file.
       01 data-rec PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-handled PIC X VALUE "N".
       PROCEDURE DIVISION.
       DECLARATIVES.
       file-err SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON data-file.
           MOVE "Y" TO ws-handled.
       END DECLARATIVES.
       main-para SECTION.
           OPEN INPUT data-file
           IF ws-handled = "N"
               DISPLAY "opened ok"
           ELSE
               DISPLAY "file error handled"
           END-IF
           STOP RUN.

