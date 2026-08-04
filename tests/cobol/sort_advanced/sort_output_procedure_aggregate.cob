*> vybe-test: cobol/sort_advanced/sort_output_procedure_aggregate
*> origin: languages/cobol/tests/cobol/test_sort_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT sf ASSIGN TO "sf.tmp".
       DATA DIVISION.
       FILE SECTION.
       SD sf.
       01 srec.
           05 dept PIC X(4).
           05 sal  PIC 9(7)V99.
       WORKING-STORAGE SECTION.
       01 ws-done      PIC X   VALUE "N".
       01 ws-prev-dept PIC X(4) VALUE SPACES.
       01 ws-dept-total PIC 9(10)V99 VALUE 0.
       PROCEDURE DIVISION.
           SORT sf ON ASCENDING KEY dept
               USING "payroll.dat"
               OUTPUT PROCEDURE IS compute-totals
           STOP RUN.
       compute-totals SECTION.
           RETURN sf AT END MOVE "Y" TO ws-done END-RETURN
           PERFORM UNTIL ws-done = "Y"
               IF dept NOT = ws-prev-dept
                   IF ws-prev-dept NOT = SPACES
                       DISPLAY ws-prev-dept ws-dept-total
                   END-IF
                   MOVE dept TO ws-prev-dept
                   MOVE 0 TO ws-dept-total
               END-IF
               ADD sal TO ws-dept-total
               RETURN sf AT END MOVE "Y" TO ws-done END-RETURN
           END-PERFORM
           IF ws-prev-dept NOT = SPACES
               DISPLAY ws-prev-dept ws-dept-total
           END-IF.

