*> vybe-test: cobol/sort_advanced/merge_with_duplicates
*> origin: languages/cobol/tests/cobol/test_sort_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT mf ASSIGN TO "mf.tmp".
       DATA DIVISION.
       FILE SECTION.
       SD mf.
       01 mrec.
           05 mk PIC 9(5).
           05 md PIC X(20).
       PROCEDURE DIVISION.
           MERGE mf
               ON ASCENDING KEY mk
               WITH DUPLICATES IN ORDER
               USING "a.dat" "b.dat" "c.dat"
               GIVING "merged.dat"
           STOP RUN.

