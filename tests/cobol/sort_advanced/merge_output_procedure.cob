*> vybe-test: cobol/sort_advanced/merge_output_procedure
*> origin: languages/cobol/tests/cobol/test_sort_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT mf  ASSIGN TO "mf.tmp".
           SELECT out ASSIGN TO "merged.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       SD mf.
       01 mrec.
           05 mk PIC X(10).
           05 md PIC X(30).
       FD out.
       01 out-rec PIC X(40).
       WORKING-STORAGE SECTION.
       01 ws-done PIC X VALUE "N".
       01 ws-count PIC 9(5) VALUE 0.
       PROCEDURE DIVISION.
           MERGE mf ON ASCENDING KEY mk
               USING "file1.dat" "file2.dat"
               OUTPUT PROCEDURE IS write-merged
           DISPLAY ws-count
           STOP RUN.
       write-merged SECTION.
           OPEN OUTPUT out
           PERFORM UNTIL ws-done = "Y"
               RETURN mf INTO mrec
                   AT END MOVE "Y" TO ws-done
                   NOT AT END
                       ADD 1 TO ws-count
                       MOVE mrec TO out-rec
                       WRITE out-rec
               END-RETURN
           END-PERFORM
           CLOSE out.

