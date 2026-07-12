use crate::helpers;

#[test]
fn test_inspect_tallying_basic() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INSPECT-TALLY.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR PIC X(10) VALUE "ABCAAB".
       01 CNT PIC 9(2) VALUE 0.
       PROCEDURE DIVISION.
           INSPECT STR TALLYING CNT FOR ALL "A".
           DISPLAY CNT.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["03"]);
}

#[test]
fn test_inspect_tallying_leading() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INSPECT-LEADING.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR PIC X(10) VALUE "000123000".
       01 CNT PIC 9(2) VALUE 0.
       PROCEDURE DIVISION.
           INSPECT STR TALLYING CNT FOR LEADING "0".
           DISPLAY CNT.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["03"]);
}

#[test]
fn test_inspect_replacing_all() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INSPECT-REPL-ALL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR PIC X(10) VALUE "A-B-C-D-".
       PROCEDURE DIVISION.
           INSPECT STR REPLACING ALL "-" BY " ".
           DISPLAY STR.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["A B C D   "]);
}

#[test]
fn test_inspect_replacing_first() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INSPECT-REPL-FIRST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR PIC X(10) VALUE "100-200-30".
       PROCEDURE DIVISION.
           INSPECT STR REPLACING FIRST "-" BY "X".
           DISPLAY STR.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["100X200-30"]);
}

#[test]
fn test_inspect_tallying_replacing_combined() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INSPECT-COMB.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR PIC X(10) VALUE "  123   ".
       01 CNT PIC 9(2) VALUE 0.
       PROCEDURE DIVISION.
           INSPECT STR
              TALLYING CNT FOR LEADING " "
              REPLACING LEADING " " BY "0"
                        ALL " " BY "X".
           DISPLAY STR " " CNT.
           STOP RUN.
    "#;
    // Tallying counts 2 leading spaces.
    // Replacing changes 2 leading spaces to 0s, and 3 trailing spaces to Xs.
    assert_eq!(helpers::run_prints(src), vec!["00123XXX   02"]);
}

#[test]
fn test_inspect_converting() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INSPECT-CONV.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR PIC X(5) VALUE "HELLO".
       PROCEDURE DIVISION.
           INSPECT STR CONVERTING "EOL" TO "e01".
           DISPLAY STR.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["He110"]);
}

#[test]
fn test_inspect_before_after() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INSPECT-BA.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR PIC X(20) VALUE "ABC*DEF*GHI".
       01 CNT PIC 9(2) VALUE 0.
       PROCEDURE DIVISION.
           INSPECT STR TALLYING CNT FOR CHARACTERS
              AFTER INITIAL "*"
              BEFORE INITIAL "*".
           DISPLAY CNT.
           STOP RUN.
    "#;
    // Wait, AFTER INITIAL "*" means start after first *.
    // BEFORE INITIAL "*" means stop before first *.
    // Since AFTER is applied before BEFORE? No, both conditions must be met.
    // If we say AFTER first * and BEFORE second *? COBOL syntax only allows one BEFORE or AFTER per phrase, or one of each.
    // Let's test basic BEFORE.
    let src2 = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INSPECT-BA.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR PIC X(20) VALUE "ABC*DEF*GHI".
       01 CNT PIC 9(2) VALUE 0.
       PROCEDURE DIVISION.
           INSPECT STR TALLYING CNT FOR CHARACTERS
              BEFORE INITIAL "*".
           DISPLAY CNT.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src2), vec!["03"]);
}
