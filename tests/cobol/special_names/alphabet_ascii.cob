*> vybe-test: cobol/special_names/alphabet_ascii
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           ALPHABET ASCII-ALPHA IS ASCII.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val PIC X VALUE "Z".
       PROCEDURE DIVISION.
           DISPLAY ws-val
           STOP RUN.

