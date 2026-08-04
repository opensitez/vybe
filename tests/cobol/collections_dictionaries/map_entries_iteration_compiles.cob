*> vybe-test: cobol/collections_dictionaries/map_entries_iteration_compiles
*> origin: languages/cobol/tests/cobol/test_collections_dictionaries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    CALL "MAP-ENTRIES".
    STOP RUN.

