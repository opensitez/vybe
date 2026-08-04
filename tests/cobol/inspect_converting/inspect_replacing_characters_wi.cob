*> vybe-test: cobol/inspect_converting/inspect_replacing_characters_with_star
*> origin: languages/cobol/tests/cobol/test_inspect_converting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "HELLO".
PROCEDURE DIVISION.
    INSPECT S REPLACING CHARACTERS BY "*".
    STOP RUN.

