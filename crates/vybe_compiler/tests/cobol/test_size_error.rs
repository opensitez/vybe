use super::helpers::compile_ok;

// ── ON SIZE ERROR / NOT ON SIZE ERROR — ADD ───────────────────

#[test] fn add_on_size_error() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-small PIC 9 VALUE 9.
       01 ws-overflow PIC X VALUE "N".
       PROCEDURE DIVISION.
           ADD 1 TO ws-small
               ON SIZE ERROR MOVE "Y" TO ws-overflow
           END-ADD
           DISPLAY ws-overflow
           STOP RUN.
"#);
}

#[test] fn add_not_on_size_error() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val     PIC 99 VALUE 5.
       01 ws-ok-flag PIC X VALUE "N".
       PROCEDURE DIVISION.
           ADD 3 TO ws-val
               NOT ON SIZE ERROR MOVE "Y" TO ws-ok-flag
           END-ADD
           DISPLAY ws-ok-flag
           STOP RUN.
"#);
}

#[test] fn add_both_size_error_branches() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-num    PIC 9 VALUE 8.
       01 ws-status PIC X VALUE SPACE.
       PROCEDURE DIVISION.
           ADD 5 TO ws-num
               ON SIZE ERROR     MOVE "O" TO ws-status
               NOT ON SIZE ERROR MOVE "K" TO ws-status
           END-ADD
           DISPLAY ws-status
           STOP RUN.
"#);
}

// ── ON SIZE ERROR — SUBTRACT ──────────────────────────────────

#[test] fn subtract_on_size_error() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-unsigned PIC 9 VALUE 0.
       01 ws-err      PIC X VALUE "N".
       PROCEDURE DIVISION.
           SUBTRACT 5 FROM ws-unsigned
               ON SIZE ERROR MOVE "Y" TO ws-err
           END-SUBTRACT
           DISPLAY ws-err
           STOP RUN.
"#);
}

#[test] fn subtract_not_on_size_error() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val PIC 99 VALUE 20.
       01 ws-ok  PIC X VALUE "N".
       PROCEDURE DIVISION.
           SUBTRACT 5 FROM ws-val
               NOT ON SIZE ERROR MOVE "Y" TO ws-ok
           END-SUBTRACT
           DISPLAY ws-ok
           STOP RUN.
"#);
}

// ── ON SIZE ERROR — MULTIPLY ──────────────────────────────────

#[test] fn multiply_on_size_error() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9 VALUE 9.
       01 ws-err    PIC X VALUE "N".
       PROCEDURE DIVISION.
           MULTIPLY 9 BY ws-result
               ON SIZE ERROR MOVE "Y" TO ws-err
           END-MULTIPLY
           DISPLAY ws-err
           STOP RUN.
"#);
}

#[test] fn multiply_giving_on_size_error() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-a      PIC 99  VALUE 50.
       01 ws-b      PIC 99  VALUE 50.
       01 ws-result PIC 999 VALUE 0.
       01 ws-err    PIC X   VALUE "N".
       PROCEDURE DIVISION.
           MULTIPLY ws-a BY ws-b GIVING ws-result
               ON SIZE ERROR MOVE "Y" TO ws-err
           END-MULTIPLY
           DISPLAY ws-err
           STOP RUN.
"#);
}

// ── ON SIZE ERROR — DIVIDE ────────────────────────────────────

#[test] fn divide_on_size_error() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-dividend PIC 9 VALUE 5.
       01 ws-divisor  PIC 9 VALUE 0.
       01 ws-err      PIC X VALUE "N".
       PROCEDURE DIVISION.
           DIVIDE ws-divisor INTO ws-dividend
               ON SIZE ERROR MOVE "Y" TO ws-err
           END-DIVIDE
           DISPLAY ws-err
           STOP RUN.
"#);
}

#[test] fn divide_giving_remainder_on_size_error() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-a      PIC 9   VALUE 7.
       01 ws-b      PIC 9   VALUE 2.
       01 ws-q      PIC 9   VALUE 0.
       01 ws-r      PIC 9   VALUE 0.
       01 ws-err    PIC X   VALUE "N".
       PROCEDURE DIVISION.
           DIVIDE ws-b INTO ws-a
               GIVING ws-q REMAINDER ws-r
               NOT ON SIZE ERROR MOVE "N" TO ws-err
           END-DIVIDE
           DISPLAY ws-q
           DISPLAY ws-r
           STOP RUN.
"#);
}

// ── ON SIZE ERROR — COMPUTE ───────────────────────────────────

#[test] fn compute_on_size_error() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9 VALUE 0.
       01 ws-err    PIC X VALUE "N".
       PROCEDURE DIVISION.
           COMPUTE ws-result = 999 * 999
               ON SIZE ERROR MOVE "Y" TO ws-err
           END-COMPUTE
           DISPLAY ws-err
           STOP RUN.
"#);
}

#[test] fn compute_not_on_size_error() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 99  VALUE 0.
       01 ws-ok     PIC X   VALUE "N".
       PROCEDURE DIVISION.
           COMPUTE ws-result = 5 + 3
               NOT ON SIZE ERROR MOVE "Y" TO ws-ok
           END-COMPUTE
           DISPLAY ws-result
           DISPLAY ws-ok
           STOP RUN.
"#);
}

#[test] fn compute_both_branches_on_size_error() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-target PIC 99 VALUE 0.
       01 ws-status PIC X  VALUE SPACE.
       PROCEDURE DIVISION.
           COMPUTE ws-target = 200 + 300
               ON SIZE ERROR     MOVE "E" TO ws-status
               NOT ON SIZE ERROR MOVE "K" TO ws-status
           END-COMPUTE
           DISPLAY ws-status
           STOP RUN.
"#);
}

// ── ON SIZE ERROR in ADD GIVING ───────────────────────────────

#[test] fn add_giving_on_size_error() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-a   PIC 999 VALUE 999.
       01 ws-b   PIC 999 VALUE 1.
       01 ws-res PIC 999 VALUE 0.
       01 ws-err PIC X   VALUE "N".
       PROCEDURE DIVISION.
           ADD ws-a ws-b GIVING ws-res
               ON SIZE ERROR MOVE "Y" TO ws-err
           END-ADD
           DISPLAY ws-err
           STOP RUN.
"#);
}

// ── Nested ON SIZE ERROR ──────────────────────────────────────

#[test] fn nested_on_size_error() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-x   PIC 9  VALUE 9.
       01 ws-y   PIC 9  VALUE 9.
       01 ws-err PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           ADD 5 TO ws-x
               ON SIZE ERROR
                   ADD 1 TO ws-err
                   ADD 5 TO ws-y
                       ON SIZE ERROR ADD 1 TO ws-err
                   END-ADD
           END-ADD
           DISPLAY ws-err
           STOP RUN.
"#);
}

// ── ON SIZE ERROR with ROUNDED ────────────────────────────────

#[test] fn compute_rounded_on_size_error() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9 VALUE 0.
       01 ws-err    PIC X VALUE "N".
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED = 8.7 + 8.7
               ON SIZE ERROR     MOVE "Y" TO ws-err
               NOT ON SIZE ERROR MOVE "N" TO ws-err
           END-COMPUTE
           DISPLAY ws-err
           STOP RUN.
"#);
}
