*> vybe-test: cobol/category_string_unstring/test_string_overflow
*> origin: languages/cobol/tests/cobol/test_category_string_unstring.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. STRING-OVERFLOW.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR PIC X(10) VALUE "0123456789".
       01 DEST PIC X(5) VALUE SPACES.
       PROCEDURE DIVISION.
           STRING STR DELIMITED BY SIZE
                  INTO DEST
                  ON OVERFLOW DISPLAY "OVERFLOW OCCURRED"
                  NOT ON OVERFLOW DISPLAY "NO OVERFLOW".
           DISPLAY DEST.
           STOP RUN.

