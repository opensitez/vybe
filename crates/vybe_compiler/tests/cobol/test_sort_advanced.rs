use super::helpers::compile_ok;

// ── SORT with DUPLICATES IN ORDER ─────────────────────────────

#[test] fn sort_duplicates_in_order_ascending() {
    compile_ok(r#"
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
"#);
}

#[test] fn sort_duplicates_in_order_descending() {
    compile_ok(r#"
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
"#);
}

#[test] fn sort_duplicates_multiple_keys() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT sort-wf ASSIGN TO "swf.tmp".
       DATA DIVISION.
       FILE SECTION.
       SD sort-wf.
       01 sort-record.
           05 dept-code  PIC X(4).
           05 emp-name   PIC X(20).
           05 salary     PIC 9(7)V99.
       PROCEDURE DIVISION.
           SORT sort-wf
               ON ASCENDING KEY dept-code
               ON ASCENDING KEY emp-name
               WITH DUPLICATES IN ORDER
               USING "employees.dat"
               GIVING "sorted-employees.dat"
           STOP RUN.
"#);
}

// ── SORT with INPUT PROCEDURE (RELEASE) ───────────────────────

#[test] fn sort_input_procedure_basic() {
    compile_ok(r#"
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
"#);
}

#[test] fn sort_input_procedure_with_filter() {
    compile_ok(r#"
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
"#);
}

#[test] fn sort_input_procedure_transform() {
    compile_ok(r#"
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
"#);
}

// ── SORT with OUTPUT PROCEDURE (RETURN) ───────────────────────

#[test] fn sort_output_procedure_basic() {
    compile_ok(r#"
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
"#);
}

#[test] fn sort_output_procedure_not_at_end() {
    compile_ok(r#"
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
"#);
}

#[test] fn sort_output_procedure_aggregate() {
    compile_ok(r#"
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
"#);
}

// ── SORT with both INPUT and OUTPUT PROCEDURE ─────────────────

#[test] fn sort_input_and_output_procedure() {
    compile_ok(r#"
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
"#);
}

// ── RELEASE statement standalone ─────────────────────────────

#[test] fn release_from_working_storage() {
    compile_ok(r#"
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
           05 sk   PIC X(5).
           05 sval PIC 9(5).
       WORKING-STORAGE SECTION.
       01 ws-key PIC X(5) VALUE "AKEY".
       01 ws-val PIC 9(5) VALUE 42.
       PROCEDURE DIVISION.
           SORT sf ON ASCENDING KEY sk
               INPUT PROCEDURE IS fill-sort
               GIVING "out.dat"
           STOP RUN.
       fill-sort SECTION.
           MOVE ws-key TO sk
           MOVE ws-val TO sval
           RELEASE srec
           MOVE "BKEY" TO sk
           MOVE 99 TO sval
           RELEASE srec FROM srec.
"#);
}

// ── RETURN statement standalone ───────────────────────────────

#[test] fn return_into_working_storage() {
    compile_ok(r#"
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
           05 sk PIC X(5).
           05 sv PIC 99.
       WORKING-STORAGE SECTION.
       01 ws-buf  PIC X(7).
       01 ws-done PIC X VALUE "N".
       01 ws-count PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           SORT sf ON ASCENDING KEY sk
               USING "in.dat"
               OUTPUT PROCEDURE IS count-recs
           DISPLAY ws-count
           STOP RUN.
       count-recs SECTION.
           PERFORM UNTIL ws-done = "Y"
               RETURN sf INTO ws-buf
                   AT END MOVE "Y" TO ws-done
                   NOT AT END ADD 1 TO ws-count
               END-RETURN
           END-PERFORM.
"#);
}

// ── MERGE with OUTPUT PROCEDURE ───────────────────────────────

#[test] fn merge_output_procedure() {
    compile_ok(r#"
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
"#);
}

#[test] fn merge_with_duplicates() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT mf ASSIGN TO "mf.tmp".
       DATA DIVISION.
       FILE SECTION.
       SD mf.
       01 mrec.
           05 mk PIC 9(5).
           05 md PIC X(20).
       PROCEDURE DIVISION.
           MERGE mf
               ON ASCENDING KEY mk
               WITH DUPLICATES IN ORDER
               USING "a.dat" "b.dat" "c.dat"
               GIVING "merged.dat"
           STOP RUN.
"#);
}
