*> vybe-test: cobol/collections_dictionaries/keyed_delete_pattern_compiles
*> origin: languages/cobol/tests/cobol/test_collections_dictionaries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 K PIC X(5) VALUE "A".
PROCEDURE DIVISION.
    CALL "MAP-DEL" USING K.
    STOP RUN.

