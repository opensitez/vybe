*> vybe-test: cobol/sort_advanced/sort_input_procedure_transform
*> origin: languages/cobol/tests/cobol/test_sort_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT srt ASSIGN TO "s.tmp".
           SELECT src ASSIGN TO "source.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       SD srt.
       01 srt-rec.
           05 srt-key  PIC X(5).
           05 srt-body PIC X(40).
       FD src.
       01 src-rec PIC X(50).
       WORKING-STORAGE SECTION.
       01 ws-eof PIC X VALUE "N".
       PROCEDURE DIVISION.
           SORT srt
               ON ASCENDING KEY srt-key
               INPUT PROCEDURE IS transform-input
               GIVING "out.dat"
           STOP RUN.
       transform-input SECTION.
           OPEN INPUT src
           READ src AT END MOVE "Y" TO ws-eof END-READ
           PERFORM UNTIL ws-eof = "Y"
               MOVE FUNCTION UPPER-CASE(src-rec(1:5)) TO srt-key
               MOVE src-rec(6:40) TO srt-body
               RELEASE srt-rec
               READ src AT END MOVE "Y" TO ws-eof END-READ
           END-PERFORM
           CLOSE src.

