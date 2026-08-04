*> vybe-test: cobol/alphabet_clause/alphabet_clause_graphic_symbols_compiles
*> origin: languages/cobol/tests/cobol/test_alphabet_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA9.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET PUNCT-SET IS "." "," ";".
PROCEDURE DIVISION.
    STOP RUN.

