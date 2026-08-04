*> vybe-test: cobol/sort_advanced/sort_output_procedure_not_at_end
*> origin: languages/cobol/tests/cobol/test_sort_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT sfile ASSIGN TO "s.tmp".
           SELECT ofile ASSIGN TO "o.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       SD sfile.
       01 srec.
           05 sk PIC 9(5).
           05 sd PIC X(30).
       FD ofile.
       01 orec PIC X(40).
       WORKING-STORAGE SECTION.
       01 ws-done PIC X VALUE "N".
       01 ws-count PIC 9(5) VALUE 0.
       PROCEDURE DIVISION.
           SORT sfile ON ASCENDING KEY sk
               USING "i.dat"
               OUTPUT PROCEDURE IS output-proc
           DISPLAY ws-count
           STOP RUN.
       output-proc SECTION.
           OPEN OUTPUT ofile
           PERFORM UNTIL ws-done = "Y"
               RETURN sfile INTO srec
                   AT END MOVE "Y" TO ws-done
                   NOT AT END
                       ADD 1 TO ws-count
                       MOVE srec TO orec
                       WRITE orec
               END-RETURN
           END-PERFORM
           CLOSE ofile.

