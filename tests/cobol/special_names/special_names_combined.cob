*> vybe-test: cobol/special_names/special_names_combined
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CURRENCY SIGN IS "$"
           DECIMAL-POINT IS COMMA
           CLASS DIGIT IS "0" THRU "9"
           ALPHABET MY-COLL IS STANDARD-1.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-price PIC $9.999,99 VALUE 1234,56.
       01 ws-ch    PIC X VALUE "7".
       PROCEDURE DIVISION.
           IF ws-ch IS DIGIT
               DISPLAY "digit"
           END-IF
           DISPLAY ws-price
           STOP RUN.

