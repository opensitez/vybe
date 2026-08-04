*> vybe-test: cobol/collating_sequence/object_computer_with_memory_and_collating_sequence_compiles
*> origin: languages/cobol/tests/cobol/test_collating_sequence.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. COLL5.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
OBJECT-COMPUTER. IBM-Z MEMORY SIZE 1024 CHARACTERS COLLATING SEQUENCE IS ALPHA5.
PROCEDURE DIVISION.
    STOP RUN.

