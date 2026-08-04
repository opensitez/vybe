*> vybe-test: cobol/string_delimited_forms/unstring_into_numeric_field
*> origin: languages/cobol/tests/cobol/test_string_delimited_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(10) VALUE "123,456".
01 F1 PIC 9(5) VALUE 0.
01 F2 PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    UNSTRING SRC DELIMITED BY "," INTO F1 F2.
    STOP RUN.

