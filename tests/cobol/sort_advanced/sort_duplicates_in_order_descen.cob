*> vybe-test: cobol/sort_advanced/sort_duplicates_in_order_descending
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
           05 sort-key PIC 9(5).
           05 sort-data PIC X(20).
       PROCEDURE DIVISION.
           SORT sort-file
               ON DESCENDING KEY sort-key
               WITH DUPLICATES IN ORDER
               USING "data.dat"
               GIVING "sorted.dat"
           STOP RUN.

