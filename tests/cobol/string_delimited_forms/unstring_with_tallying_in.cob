*> vybe-test: cobol/string_delimited_forms/unstring_with_tallying_in
*> origin: languages/cobol/tests/cobol/test_string_delimited_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(20) VALUE "A,B,C,D".
01 F1 PIC X(5).
01 F2 PIC X(5).
01 F3 PIC X(5).
01 CNT PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    UNSTRING SRC DELIMITED BY "," INTO F1 F2 F3 TALLYING IN CNT.
    STOP RUN.

