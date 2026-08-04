*> vybe-test: cobol/regex_text_features/inspect_replace_letters_compiles
*> origin: languages/cobol/tests/cobol/test_regex_text_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(30).
PROCEDURE DIVISION.
    INSPECT S REPLACING ALL "A" BY "B".
    STOP RUN.

