*> vybe-test: cobol/collating_sequence/collating_sequence_special_names_with_alphabet_compiles
*> origin: languages/cobol/tests/cobol/test_collating_sequence.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. COLL4.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET MY-ALPHA IS STANDARD-1.
    COLLATING SEQUENCE IS MY-ALPHA.
PROCEDURE DIVISION.
    STOP RUN.

