*> vybe-test: cobol/string_delimited_forms/unstring_delimiter_in_field_compiles
*> origin: languages/cobol/tests/cobol/test_string_delimited_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(30) VALUE "a=1;b=2".
01 K1 PIC X(5).
01 V1 PIC X(5).
01 K2 PIC X(5).
01 V2 PIC X(5).
PROCEDURE DIVISION.
    UNSTRING SRC DELIMITED BY "=" OR ";" INTO K1 V1 K2 V2.
    STOP RUN.

