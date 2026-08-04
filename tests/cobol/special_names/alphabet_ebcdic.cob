*> vybe-test: cobol/special_names/alphabet_ebcdic
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           ALPHABET EBCDIC-ALPHA IS EBCDIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-char PIC X VALUE "A".
       PROCEDURE DIVISION.
           DISPLAY ws-char
           STOP RUN.

