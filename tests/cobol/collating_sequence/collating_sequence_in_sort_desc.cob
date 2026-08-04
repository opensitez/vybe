*> vybe-test: cobol/collating_sequence/collating_sequence_in_sort_descending_compiles
*> origin: languages/cobol/tests/cobol/test_collating_sequence.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. COLL6.
DATA DIVISION.
SD SORT-FILE.
01 SREC PIC X(10).
WORKING-STORAGE SECTION.
01 KEY PIC X(5).
PROCEDURE DIVISION.
    SORT SORT-FILE ON DESCENDING KEY KEY COLLATING SEQUENCE IS ALPHA6.
    STOP RUN.

