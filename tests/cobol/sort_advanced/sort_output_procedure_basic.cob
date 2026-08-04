*> vybe-test: cobol/sort_advanced/sort_output_procedure_basic
*> origin: languages/cobol/tests/cobol/test_sort_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT sort-file  ASSIGN TO "sort.tmp".
           SELECT out-file   ASSIGN TO "output.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       SD sort-file.
       01 sort-rec.
           05 s-key  PIC X(10).
           05 s-body PIC X(70).
       FD out-file.
       01 out-rec PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-end-sort PIC X VALUE "N".
       PROCEDURE DIVISION.
           SORT sort-file
               ON ASCENDING KEY s-key
               USING "input.dat"
               OUTPUT PROCEDURE IS process-output
           STOP RUN.
       process-output SECTION.
           OPEN OUTPUT out-file
           RETURN sort-file INTO out-rec
               AT END MOVE "Y" TO ws-end-sort
           END-RETURN
           PERFORM UNTIL ws-end-sort = "Y"
               WRITE out-rec
               RETURN sort-file INTO out-rec
                   AT END MOVE "Y" TO ws-end-sort
               END-RETURN
           END-PERFORM
           CLOSE out-file.

