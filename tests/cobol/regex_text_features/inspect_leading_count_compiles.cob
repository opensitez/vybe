*> vybe-test: cobol/regex_text_features/inspect_leading_count_compiles
*> origin: languages/cobol/tests/cobol/test_regex_text_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(30).
01 C PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    INSPECT S TALLYING C FOR LEADING "0".
    STOP RUN.

