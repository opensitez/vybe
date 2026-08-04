*> vybe-test: cobol/category_string_unstring/test_unstring_pointer_overflow
*> origin: languages/cobol/tests/cobol/test_category_string_unstring.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. UNSTRING-PTR-OVF.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 SRC PIC X(10) VALUE "A,B,C,D,E".
       01 OUT-1 PIC X(2).
       01 OUT-2 PIC X(2).
       01 PTR PIC 9(2) VALUE 1.
       PROCEDURE DIVISION.
           UNSTRING SRC DELIMITED BY ","
              INTO OUT-1 OUT-2
              WITH POINTER PTR
              ON OVERFLOW DISPLAY "OVERFLOW"
           END-UNSTRING.
           DISPLAY PTR.
           STOP RUN.

