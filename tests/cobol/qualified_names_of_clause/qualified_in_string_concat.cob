*> vybe-test: cobol/qualified_names_of_clause/qualified_in_string_concat
*> origin: languages/cobol/tests/cobol/test_qualified_names_of_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC-A.
   05 WORD PIC X(5) VALUE "HELLO".
01 SRC-B.
   05 WORD PIC X(5) VALUE "WORLD".
01 R PIC X(15) VALUE SPACES.
PROCEDURE DIVISION.
    STRING WORD OF SRC-A DELIMITED BY SPACE " " DELIMITED BY SIZE WORD OF SRC-B DELIMITED BY SPACE INTO R.
    STOP RUN.

