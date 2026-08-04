*> vybe-test: cobol/regex_text_features/inspect_first_replacement_compiles
*> origin: languages/cobol/tests/cobol/test_regex_text_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(30).
PROCEDURE DIVISION.
    INSPECT S REPLACING FIRST "A" BY "Z".
    STOP RUN.

