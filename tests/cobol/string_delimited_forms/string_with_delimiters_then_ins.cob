*> vybe-test: cobol/string_delimited_forms/string_with_delimiters_then_inspect
*> origin: languages/cobol/tests/cobol/test_string_delimited_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 NAME PIC X(10) VALUE "JOHN      ".
01 R PIC X(20) VALUE SPACES.
01 C PIC 9(2) VALUE 0.
PROCEDURE DIVISION.
    STRING NAME DELIMITED BY SPACE INTO R.
    INSPECT R TALLYING C FOR ALL "O".
    STOP RUN.

