*> vybe-test: cobol/figurative_constants/low_values_as_minimum_key
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-search-key PIC X(10).
       01 ws-result     PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE LOW-VALUES TO ws-search-key
           IF ws-search-key < "AAAA"
               MOVE "Y" TO ws-result
           END-IF
           DISPLAY ws-result
           STOP RUN.

