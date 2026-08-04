*> vybe-test: cobol/sort_advanced/sort_input_procedure_with_filter
*> origin: languages/cobol/tests/cobol/test_sort_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT sort-wf  ASSIGN TO "sort.tmp".
           SELECT raw-file ASSIGN TO "raw.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       SD sort-wf.
       01 sort-record.
           05 sr-score PIC 999.
           05 sr-name  PIC X(30).
       FD raw-file.
       01 raw-rec.
           05 rr-score PIC 999.
           05 rr-name  PIC X(30).
       WORKING-STORAGE SECTION.
       01 ws-eof PIC X VALUE "N".
       PROCEDURE DIVISION.
           SORT sort-wf
               ON DESCENDING KEY sr-score
               INPUT PROCEDURE IS filter-and-release
               GIVING "high-scores.dat"
           STOP RUN.
       filter-and-release SECTION.
           OPEN INPUT raw-file
           READ raw-file AT END MOVE "Y" TO ws-eof END-READ
           PERFORM UNTIL ws-eof = "Y"
               IF rr-score >= 60
                   MOVE rr-score TO sr-score
                   MOVE rr-name  TO sr-name
                   RELEASE sort-record
               END-IF
               READ raw-file AT END MOVE "Y" TO ws-eof END-READ
           END-PERFORM
           CLOSE raw-file.

