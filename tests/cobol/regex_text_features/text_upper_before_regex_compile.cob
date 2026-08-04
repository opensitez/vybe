*> vybe-test: cobol/regex_text_features/text_upper_before_regex_compiles
*> origin: languages/cobol/tests/cobol/test_regex_text_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(30).
01 O PIC X(30).
PROCEDURE DIVISION.
    MOVE FUNCTION UPPER-CASE(S) TO O.
    STOP RUN.

