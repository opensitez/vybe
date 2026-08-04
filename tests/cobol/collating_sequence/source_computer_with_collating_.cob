*> vybe-test: cobol/collating_sequence/source_computer_with_collating_context_compiles
*> origin: languages/cobol/tests/cobol/test_collating_sequence.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. COLL9.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SOURCE-COMPUTER. IBM-Z.
SPECIAL-NAMES.
    COLLATING SEQUENCE IS ALPHA9.
PROCEDURE DIVISION.
    STOP RUN.

