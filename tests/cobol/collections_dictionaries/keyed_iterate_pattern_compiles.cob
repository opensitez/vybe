*> vybe-test: cobol/collections_dictionaries/keyed_iterate_pattern_compiles
*> origin: languages/cobol/tests/cobol/test_collections_dictionaries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 1.
PROCEDURE DIVISION.
    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 5
        CALL "MAP-NEXT" USING I
    END-PERFORM.
    STOP RUN.

