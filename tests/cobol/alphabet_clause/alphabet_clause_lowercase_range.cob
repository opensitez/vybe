*> vybe-test: cobol/alphabet_clause/alphabet_clause_lowercase_range_compiles
*> origin: languages/cobol/tests/cobol/test_alphabet_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA5.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET LOWER-SET IS "a" THRU "z".
PROCEDURE DIVISION.
    STOP RUN.

