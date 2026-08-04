*> vybe-test: cobol/strings/unstring_comma
*> origin: languages/cobol/tests/cobol/test_strings.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(30) VALUE "A,B,C".
01 F1 PIC X(10).
01 F2 PIC X(10).
01 F3 PIC X(10).
PROCEDURE DIVISION.
    UNSTRING SRC DELIMITED BY "," INTO F1 F2 F3.
    STOP RUN.

