*> vybe-test: cobol/collating_sequence/sort_collating_sequence_compiles
*> origin: languages/cobol/tests/cobol/test_collating_sequence.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. COLL3.
DATA DIVISION.
SD SORT-FILE.
01 SREC PIC X(10).
WORKING-STORAGE SECTION.
01 KEY PIC X(5).
PROCEDURE DIVISION.
    SORT SORT-FILE ON ASCENDING KEY KEY COLLATING SEQUENCE IS ALPHA3.
    STOP RUN.

