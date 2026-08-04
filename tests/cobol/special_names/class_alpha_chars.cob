*> vybe-test: cobol/special_names/class_alpha_chars
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CLASS UPPER-ALPHA IS "A" THRU "Z"
           CLASS LOWER-ALPHA IS "a" THRU "z".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-ch PIC X VALUE "M".
       PROCEDURE DIVISION.
           EVALUATE TRUE
               WHEN ws-ch IS UPPER-ALPHA DISPLAY "upper"
               WHEN ws-ch IS LOWER-ALPHA DISPLAY "lower"
               WHEN OTHER                DISPLAY "other"
           END-EVALUATE
           STOP RUN.

