*> vybe-test: cobol/strings/unstring_space
*> origin: languages/cobol/tests/cobol/test_strings.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(30) VALUE "Hello World Cobol".
01 W1 PIC X(10).
01 W2 PIC X(10).
01 W3 PIC X(10).
PROCEDURE DIVISION.
    UNSTRING SRC DELIMITED BY " " INTO W1 W2 W3.
    STOP RUN.

