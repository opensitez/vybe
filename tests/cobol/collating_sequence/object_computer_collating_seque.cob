*> vybe-test: cobol/collating_sequence/object_computer_collating_sequence_compiles
*> origin: languages/cobol/tests/cobol/test_collating_sequence.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. COLL2.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
OBJECT-COMPUTER. IBM-Z COLLATING SEQUENCE IS ALPHA2.
PROCEDURE DIVISION.
    STOP RUN.

