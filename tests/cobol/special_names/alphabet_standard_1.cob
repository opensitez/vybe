*> vybe-test: cobol/special_names/alphabet_standard_1
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           ALPHABET STD-ALPHA IS STANDARD-1.
       PROCEDURE DIVISION.
           DISPLAY "ok"
           STOP RUN.

