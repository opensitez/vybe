*> vybe-test: cobol/string_delimited_forms/unstring_with_count_in
*> origin: languages/cobol/tests/cobol/test_string_delimited_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(15) VALUE "ABC,DEF".
01 F1 PIC X(5).
01 F2 PIC X(5).
01 C1 PIC 9(3) VALUE 0.
01 C2 PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    UNSTRING SRC DELIMITED BY "," INTO F1 COUNT IN C1 F2 COUNT IN C2.
    STOP RUN.

