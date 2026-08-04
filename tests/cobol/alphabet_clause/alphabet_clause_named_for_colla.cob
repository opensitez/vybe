*> vybe-test: cobol/alphabet_clause/alphabet_clause_named_for_collating_sequence_compiles
*> origin: languages/cobol/tests/cobol/test_alphabet_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA10.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET MY-ALPHA IS STANDARD-1.
    COLLATING SEQUENCE IS MY-ALPHA.
PROCEDURE DIVISION.
    STOP RUN.

