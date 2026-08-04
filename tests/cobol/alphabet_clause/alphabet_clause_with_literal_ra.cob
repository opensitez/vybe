*> vybe-test: cobol/alphabet_clause/alphabet_clause_with_literal_range_compiles
*> origin: languages/cobol/tests/cobol/test_alphabet_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA2.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET ALPHA-2 IS "A" THRU "Z".
PROCEDURE DIVISION.
    STOP RUN.

