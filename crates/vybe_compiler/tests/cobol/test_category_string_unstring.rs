use crate::helpers;

#[test]
fn test_string_basic_concatenation() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. STRING-BASIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR-1 PIC X(3) VALUE "ABC".
       01 STR-2 PIC X(3) VALUE "DEF".
       01 DEST PIC X(10) VALUE SPACES.
       PROCEDURE DIVISION.
           STRING STR-1 DELIMITED BY SIZE
                  STR-2 DELIMITED BY SIZE
                  INTO DEST.
           DISPLAY DEST.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["ABCDEF    "]);
}

#[test]
fn test_string_delimited_by_char() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. STRING-DELIM-CHAR.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR-1 PIC X(10) VALUE "HELLO*FOO".
       01 STR-2 PIC X(10) VALUE "WORLD*BAR".
       01 DEST PIC X(20) VALUE SPACES.
       PROCEDURE DIVISION.
           STRING STR-1 DELIMITED BY "*"
                  " " DELIMITED BY SIZE
                  STR-2 DELIMITED BY "*"
                  INTO DEST.
           DISPLAY DEST.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["HELLO WORLD         "]);
}

#[test]
fn test_string_with_pointer() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. STRING-PTR.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR-1 PIC X(5) VALUE "COBOL".
       01 DEST PIC X(10) VALUE SPACES.
       01 PTR PIC 9(2) VALUE 3.
       PROCEDURE DIVISION.
           STRING STR-1 DELIMITED BY SIZE
                  INTO DEST WITH POINTER PTR.
           DISPLAY DEST.
           DISPLAY PTR.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["  COBOL   ", "08"]);
}

#[test]
fn test_string_overflow() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. STRING-OVERFLOW.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR PIC X(10) VALUE "0123456789".
       01 DEST PIC X(5) VALUE SPACES.
       PROCEDURE DIVISION.
           STRING STR DELIMITED BY SIZE
                  INTO DEST
                  ON OVERFLOW DISPLAY "OVERFLOW OCCURRED"
                  NOT ON OVERFLOW DISPLAY "NO OVERFLOW".
           DISPLAY DEST.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["OVERFLOW OCCURRED", "01234"]);
}

#[test]
fn test_unstring_basic() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. UNSTRING-BASIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 SRC PIC X(20) VALUE "PART1,PART2".
       01 OUT-1 PIC X(10).
       01 OUT-2 PIC X(10).
       PROCEDURE DIVISION.
           UNSTRING SRC DELIMITED BY ","
              INTO OUT-1 OUT-2.
           DISPLAY OUT-1 "|" OUT-2.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["PART1     |PART2     "]);
}

#[test]
fn test_unstring_multiple_delimiters() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. UNSTRING-MULTI.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 SRC PIC X(20) VALUE "A,B;C D".
       01 OUT-1 PIC X(2).
       01 OUT-2 PIC X(2).
       01 OUT-3 PIC X(2).
       01 OUT-4 PIC X(2).
       PROCEDURE DIVISION.
           UNSTRING SRC DELIMITED BY "," OR ";" OR " "
              INTO OUT-1 OUT-2 OUT-3 OUT-4.
           DISPLAY OUT-1 "|" OUT-2 "|" OUT-3 "|" OUT-4.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["A |B |C |D "]);
}

#[test]
fn test_unstring_delimited_by_all() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. UNSTRING-ALL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 SRC PIC X(20) VALUE "A,,,,,B".
       01 OUT-1 PIC X(5).
       01 OUT-2 PIC X(5).
       PROCEDURE DIVISION.
           UNSTRING SRC DELIMITED BY ALL ","
              INTO OUT-1 OUT-2.
           DISPLAY OUT-1 "|" OUT-2.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["A    |B    "]);
}

#[test]
fn test_unstring_tallying_count_in() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. UNSTRING-TALLY.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 SRC PIC X(20) VALUE "APPLE,BANANA".
       01 OUT-1 PIC X(10).
       01 CNT-1 PIC 9(2) VALUE 0.
       01 OUT-2 PIC X(10).
       01 CNT-2 PIC 9(2) VALUE 0.
       01 TALLY PIC 9(2) VALUE 0.
       PROCEDURE DIVISION.
           UNSTRING SRC DELIMITED BY ","
              INTO OUT-1 COUNT IN CNT-1
                   OUT-2 COUNT IN CNT-2
              TALLYING IN TALLY.
           DISPLAY CNT-1 " " CNT-2 " " TALLY.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["05 06 02"]);
}

#[test]
fn test_unstring_pointer_overflow() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. UNSTRING-PTR-OVF.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 SRC PIC X(10) VALUE "A,B,C,D,E".
       01 OUT-1 PIC X(2).
       01 OUT-2 PIC X(2).
       01 PTR PIC 9(2) VALUE 1.
       PROCEDURE DIVISION.
           UNSTRING SRC DELIMITED BY ","
              INTO OUT-1 OUT-2
              WITH POINTER PTR
              ON OVERFLOW DISPLAY "OVERFLOW"
           END-UNSTRING.
           DISPLAY PTR.
           STOP RUN.
    "#;
    // Unstrings A into OUT-1, B into OUT-2. C,D,E are left over, causing overflow.
    // PTR starts at 1, goes past A, (2), goes past B, (4). Wait, A is pos 1. ',' is pos 2. 'B' is pos 3. ',' is pos 4. PTR becomes 5.
    assert_eq!(helpers::run_prints(src), vec!["OVERFLOW", "05"]);
}
