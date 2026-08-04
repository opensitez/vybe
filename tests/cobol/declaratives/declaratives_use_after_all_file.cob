*> vybe-test: cobol/declaratives/declaratives_use_after_all_files
*> origin: languages/cobol/tests/cobol/test_declaratives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-count PIC 99 VALUE 0.
       PROCEDURE DIVISION.
       DECLARATIVES.
       all-file-errors SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON ALL.
           ADD 1 TO ws-count.
       END DECLARATIVES.
       main-section SECTION.
           DISPLAY ws-count
           STOP RUN.

