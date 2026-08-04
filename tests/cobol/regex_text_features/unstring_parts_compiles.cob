*> vybe-test: cobol/regex_text_features/unstring_parts_compiles
*> origin: languages/cobol/tests/cobol/test_regex_text_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(20).
01 A PIC X(10).
01 B PIC X(10).
PROCEDURE DIVISION.
    UNSTRING S DELIMITED BY "," INTO A B.
    STOP RUN.

