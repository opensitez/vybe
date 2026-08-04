*> vybe-test: cobol/regex_text_features/reference_modification_compiles
*> origin: languages/cobol/tests/cobol/test_regex_text_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(30).
01 O PIC X(10).
PROCEDURE DIVISION.
    MOVE S(1:10) TO O.
    STOP RUN.

