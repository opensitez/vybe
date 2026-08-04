*> vybe-test: cobol/level_66_78/renames_in_move
*> origin: languages/cobol/tests/cobol/test_level_66_78.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 source-rec.
           05 src-a PIC X(5).
           05 src-b PIC X(5).
           05 src-c PIC X(5).
       01 target-rec.
           05 tgt-data PIC X(15).
       66 src-ab RENAMES src-a THRU src-b.
       PROCEDURE DIVISION.
           MOVE "AAAA " TO src-a
           MOVE "BBBBB" TO src-b
           MOVE src-ab TO tgt-data
           DISPLAY tgt-data
           STOP RUN.

