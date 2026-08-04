*> vybe-test: cobol/special_names/alphabet_user_defined
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           ALPHABET MY-ALPHA IS "A" "B" "C" "D" THRU "Z".
       PROCEDURE DIVISION.
           DISPLAY "ok"
           STOP RUN.

