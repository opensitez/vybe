*> vybe-test: cobol/qualified_names_of_clause/qualified_string_delimited_field
*> origin: languages/cobol/tests/cobol/test_qualified_names_of_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 PART-A.
   05 LABEL PIC X(10) VALUE "HELLO     ".
01 PART-B.
   05 LABEL PIC X(10) VALUE "WORLD     ".
01 RESULT PIC X(25) VALUE SPACES.
PROCEDURE DIVISION.
    STRING LABEL OF PART-A DELIMITED BY SPACE " " DELIMITED BY SIZE LABEL OF PART-B DELIMITED BY SPACE INTO RESULT.
    STOP RUN.

