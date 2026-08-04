*> vybe-test: cobol/alphabet_clause/alphabet_clause_digits_range_compiles
*> origin: languages/cobol/tests/cobol/test_alphabet_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA6.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET DIGIT-SET IS "0" THRU "9".
PROCEDURE DIVISION.
    STOP RUN.

