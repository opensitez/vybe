*> vybe-test: cobol/unstring_advanced/test_unstring_all_delimiters
*> origin: languages/cobol/tests/cobol/test_unstring_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SRC PIC X(10) VALUE "A,,BB,,CCC".
01 WS-F1 PIC X(3).
01 WS-F2 PIC X(3).
PROCEDURE DIVISION.

    UNSTRING WS-SRC DELIMITED BY ALL "," INTO WS-F1 WS-F2.
    STOP RUN.

