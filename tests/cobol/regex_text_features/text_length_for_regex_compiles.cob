*> vybe-test: cobol/regex_text_features/text_length_for_regex_compiles
*> origin: languages/cobol/tests/cobol/test_regex_text_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(30).
01 L PIC 9(3).
PROCEDURE DIVISION.
    MOVE FUNCTION LENGTH(S) TO L.
    STOP RUN.

