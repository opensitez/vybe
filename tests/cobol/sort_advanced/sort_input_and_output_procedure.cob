*> vybe-test: cobol/sort_advanced/sort_input_and_output_procedure
*> origin: languages/cobol/tests/cobol/test_sort_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT sf ASSIGN TO "sf.tmp".
           SELECT src ASSIGN TO "src.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
           SELECT dst ASSIGN TO "dst.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       SD sf.
       01 srec.
           05 sk PIC X(10).
           05 sd PIC X(30).
       FD src.
       01 src-rec PIC X(40).
       FD dst.
       01 dst-rec PIC X(40).
       WORKING-STORAGE SECTION.
       01 ws-eof  PIC X VALUE "N".
       01 ws-done PIC X VALUE "N".
       PROCEDURE DIVISION.
           SORT sf
               ON ASCENDING KEY sk
               INPUT PROCEDURE IS load-data
               OUTPUT PROCEDURE IS save-data
           STOP RUN.
       load-data SECTION.
           OPEN INPUT src
           READ src AT END MOVE "Y" TO ws-eof END-READ
           PERFORM UNTIL ws-eof = "Y"
               MOVE src-rec(1:10) TO sk
               MOVE src-rec(11:30) TO sd
               RELEASE srec
               READ src AT END MOVE "Y" TO ws-eof END-READ
           END-PERFORM
           CLOSE src.
       save-data SECTION.
           OPEN OUTPUT dst
           RETURN sf AT END MOVE "Y" TO ws-done END-RETURN
           PERFORM UNTIL ws-done = "Y"
               MOVE srec TO dst-rec
               WRITE dst-rec
               RETURN sf AT END MOVE "Y" TO ws-done END-RETURN
           END-PERFORM
           CLOSE dst.

