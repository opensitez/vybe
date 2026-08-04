*> vybe-test: cobol/string_delimited_forms/unstring_all_delimiter_collapses
*> origin: languages/cobol/tests/cobol/test_string_delimited_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(20) VALUE "A,,B".
01 F1 PIC X(5).
01 F2 PIC X(5).
PROCEDURE DIVISION.
    UNSTRING SRC DELIMITED BY ALL "," INTO F1 F2.
    STOP RUN.

