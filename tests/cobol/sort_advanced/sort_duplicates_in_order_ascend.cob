*> vybe-test: cobol/sort_advanced/sort_duplicates_in_order_ascending
*> origin: languages/cobol/tests/cobol/test_sort_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT sort-file ASSIGN TO "sort.tmp".
       DATA DIVISION.
       FILE SECTION.
       SD sort-file.
       01 sort-rec.
           05 sort-key   PIC X(10).
           05 sort-seq   PIC 99.
       WORKING-STORAGE SECTION.
       01 ws-done PIC X VALUE "N".
       PROCEDURE DIVISION.
           SORT sort-file
               ON ASCENDING KEY sort-key
               WITH DUPLICATES IN ORDER
               USING "input.dat"
               GIVING "output.dat"
           STOP RUN.

