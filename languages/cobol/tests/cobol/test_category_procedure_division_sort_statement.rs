use super::helpers::{compile_ok, run_prints};

#[test]
fn sort_statement_with_input_and_output_procedures_runtime() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. SRT1.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT S ASSIGN TO "SRT1".
DATA DIVISION.
FILE SECTION.
SD S.
01 R.
    05 K PIC X(1).
    05 V PIC X(3).
PROCEDURE DIVISION.
    SORT S
        ON ASCENDING KEY K
        INPUT PROCEDURE IS SRT-IN
        OUTPUT PROCEDURE IS SRT-OUT.
    STOP RUN.
SRT-IN SECTION.
    MOVE "C" TO R
    RELEASE R
    MOVE "A" TO R
    RELEASE R
    MOVE "B" TO R
    RELEASE R.
SRT-OUT SECTION.
    RETURN S AT END DISPLAY "DONE" END-RETURN
    PERFORM UNTIL FALSE
        RETURN S
            AT END DISPLAY "DONE"
            GO TO SRT-OUT-DONE
            NOT AT END DISPLAY K
        END-RETURN
    END-PERFORM
SRT-OUT-DONE.
    EXIT.
"#,
    );
    assert_eq!(out, vec!["A", "B", "C", "DONE"]);
}

#[test]
fn sort_statement_descending_key_runtime() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. SRT2.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT S ASSIGN TO "SRT2".
DATA DIVISION.
FILE SECTION.
SD S.
01 R.
    05 K PIC X(1).
    05 V PIC X(3).
PROCEDURE DIVISION.
    SORT S
        ON DESCENDING KEY K
        INPUT PROCEDURE SRT-IN
        OUTPUT PROCEDURE SRT-OUT.
    STOP RUN.
SRT-IN SECTION.
    MOVE "A" TO R
    RELEASE R
    MOVE "C" TO R
    RELEASE R
    MOVE "B" TO R
    RELEASE R.
SRT-OUT SECTION.
    RETURN S AT END GO TO SRT-OUT-DONE END-RETURN
    DISPLAY K
    GO TO SRT-OUT.
SRT-OUT-DONE.
    DISPLAY "DONE".
"#,
    );
    assert_eq!(out, vec!["C", "B", "A", "DONE"]);
}

#[test]
fn sort_statement_multiple_keys_runtime() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. SRT3.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT S ASSIGN TO "SRT3".
DATA DIVISION.
FILE SECTION.
SD S.
01 R.
    05 K1 PIC X(1).
    05 K2 PIC 9.
    05 V  PIC X(2).
PROCEDURE DIVISION.
    SORT S
        ON ASCENDING KEY K1
        ON DESCENDING KEY K2
        INPUT PROCEDURE SRT-IN
        OUTPUT PROCEDURE SRT-OUT.
    STOP RUN.
SRT-IN SECTION.
    MOVE "A2" TO R
    RELEASE R
    MOVE "A9" TO R
    RELEASE R
    MOVE "B1" TO R
    RELEASE R.
SRT-OUT SECTION.
    RETURN S
        AT END DISPLAY "DONE"
        NOT AT END
            DISPLAY K1
            DISPLAY K2
        END-RETURN.
"#,
    );
    assert_eq!(out, vec!["A", "9", "A", "2", "B", "1", "DONE"]);
}

#[test]
fn sort_statement_with_duplicate_control_is_accepted() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SRT4.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT S ASSIGN TO "SRT4".
       DATA DIVISION.
       FILE SECTION.
       SD S.
       01 R.
           05 K PIC X(1).
           05 V PIC X(3).
       PROCEDURE DIVISION.
           SORT S
               ON ASCENDING KEY K
               WITH DUPLICATES IN ORDER
               INPUT PROCEDURE SRT-IN
               OUTPUT PROCEDURE SRT-OUT.
           STOP RUN.
       SRT-IN.
       SRT-OUT.
"#,
    );
}

#[test]
fn sort_statement_without_is_keyword_is_accepted() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SRT5.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT S ASSIGN TO "SRT5".
       DATA DIVISION.
       FILE SECTION.
       SD S.
       01 R.
           05 K PIC X(1).
           05 V PIC X(3).
       PROCEDURE DIVISION.
           SORT S
               ON ASCENDING KEY K
               INPUT PROCEDURE SRT-IN
               OUTPUT PROCEDURE SRT-OUT.
           STOP RUN.
       SRT-IN.
       SRT-OUT.
"#,
    );
}

#[test]
fn sort_statement_runtime_giving_file_compiles() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SRT6.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT S ASSIGN TO "SRT6".
       DATA DIVISION.
       FILE SECTION.
        SD S.
        01 R.
            05 K PIC X(1).
            05 V PIC X(3).
        PROCEDURE DIVISION.
            SORT S
                ON ASCENDING KEY K
                USING "input.dat"
                GIVING "out.dat"
            STOP RUN.
"#,
    );
}

#[test]
fn sort_statement_output_procedure_without_is_keyword() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. SRT7.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT S ASSIGN TO "SRT7".
DATA DIVISION.
FILE SECTION.
SD S.
01 R.
    05 K PIC X(1).
    05 V PIC X(3).
PROCEDURE DIVISION.
    SORT S
        ON ASCENDING KEY K
        INPUT PROCEDURE SRT-IN
        OUTPUT PROCEDURE SRT-OUT.
    STOP RUN.
SRT-IN.
    MOVE "X" TO R
    RELEASE R.
SRT-OUT.
    RETURN S
        AT END DISPLAY "DONE"
        NOT AT END DISPLAY K
        END-RETURN.
"#,
    );
    assert_eq!(out, vec!["X", "DONE"]);
}
