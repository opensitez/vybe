*> vybe-test: cobol/collections_dictionaries/map_value_update_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that does not exist in this run unit. cobc
*> compiles it and then aborts — `libcob: error: module not found` — so
*> "runs and exits 0" is not a property this source has under any COBOL.
*> What it CAN assert is the one its name claims: that it compiles.
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

