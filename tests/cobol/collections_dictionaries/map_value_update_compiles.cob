*> vybe-test: cobol/collections_dictionaries/map_value_update_compiles
*> origin: languages/cobol/tests/cobol/test_collections_dictionaries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 K PIC X(5) VALUE "A".
01 V PIC X(10) VALUE "TWO".
PROCEDURE DIVISION.
    CALL "MAP-PUT" USING K V.
    STOP RUN.

