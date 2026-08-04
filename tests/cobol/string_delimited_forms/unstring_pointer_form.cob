*> vybe-test: cobol/string_delimited_forms/unstring_pointer_form
*> origin: languages/cobol/tests/cobol/test_string_delimited_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(20) VALUE "HELLO WORLD".
01 F1 PIC X(8).
01 PTR PIC 9(3) VALUE 1.
PROCEDURE DIVISION.
    UNSTRING SRC DELIMITED BY SPACE INTO F1 WITH POINTER PTR.
    STOP RUN.

