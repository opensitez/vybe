*> vybe-test: cobol/sort_advanced/sort_input_procedure_basic
*> origin: languages/cobol/tests/cobol/test_sort_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT sort-file ASSIGN TO "sort.tmp".
           SELECT in-file   ASSIGN TO "input.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       SD sort-file.
       01 sort-rec.
           05 s-key  PIC X(10).
           05 s-data PIC X(70).
       FD in-file.
       01 in-rec PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-eof PIC X VALUE "N".
       PROCEDURE DIVISION.
           SORT sort-file
               ON ASCENDING KEY s-key
               INPUT PROCEDURE IS sort-input
               GIVING "output.dat"
           STOP RUN.
       sort-input SECTION.
           OPEN INPUT in-file
           READ in-file
               AT END MOVE "Y" TO ws-eof
           END-READ
           PERFORM UNTIL ws-eof = "Y"
               MOVE in-rec(1:10) TO s-key
               MOVE in-rec(11:70) TO s-data
               RELEASE sort-rec
               READ in-file
                   AT END MOVE "Y" TO ws-eof
               END-READ
           END-PERFORM
           CLOSE in-file.

