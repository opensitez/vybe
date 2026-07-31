use crate::helpers;

#[test]
fn test_table_basic_occurs() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-OCCURS.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ARR.
          05 ELEM OCCURS 5 TIMES PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           MOVE 7 TO ELEM(3).
           DISPLAY ELEM(3).
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["7"]);
}

#[test]
fn test_table_multidimensional() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-MULTI.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 MATRIX.
          05 ROW OCCURS 3 TIMES.
             10 COL OCCURS 3 TIMES PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           MOVE 9 TO COL(2, 3).
           DISPLAY COL(2, 3).
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["9"]);
}

#[test]
fn test_table_indexed_by() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-INDEX.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ARR.
          05 ELEM OCCURS 5 TIMES INDEXED BY IDX PIC X(2).
       PROCEDURE DIVISION.
           SET IDX TO 2.
           MOVE "AB" TO ELEM(IDX).
           DISPLAY ELEM(IDX).
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["AB"]);
}

#[test]
fn test_table_occurs_depending_on() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-ODO.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ARR-LEN PIC 9 VALUE 2.
       01 ARR.
          05 ELEM OCCURS 1 TO 5 TIMES DEPENDING ON ARR-LEN PIC X.
       PROCEDURE DIVISION.
           MOVE "A" TO ELEM(1).
           MOVE "B" TO ELEM(2).
           DISPLAY ARR.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["AB"]);
}

#[test]
fn test_table_relative_indexing() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-REL-IDX.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ARR.
          05 ELEM OCCURS 5 TIMES INDEXED BY IDX PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           SET IDX TO 2.
           MOVE 5 TO ELEM(IDX + 1).
           MOVE 4 TO ELEM(IDX - 1).
           DISPLAY ELEM(1) ELEM(3).
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["45"]);
}

#[test]
fn test_table_set_up_down() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-SET-MATH.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ARR.
          05 ELEM OCCURS 5 TIMES INDEXED BY IDX PIC 9.
       PROCEDURE DIVISION.
           SET IDX TO 1.
           SET IDX UP BY 2.
           SET IDX DOWN BY 1.
           MOVE 8 TO ELEM(IDX).
           DISPLAY ELEM(2).
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["8"]);
}

#[test]
fn test_table_initialization() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-INIT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ARR.
          05 ELEM OCCURS 3 TIMES PIC X VALUE "*".
       PROCEDURE DIVISION.
           DISPLAY ARR.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["***"]);
}

#[test]
fn test_table_search_runtime_found() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-SRCH.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS.
          05 ELEM OCCURS 4 TIMES INDEXED BY I PIC 9(2).
       PROCEDURE DIVISION.
           MOVE 10 TO ELEM(1).
           MOVE 20 TO ELEM(2).
           MOVE 30 TO ELEM(3).
           SEARCH WS
               WHEN ELEM(I) = 20
                   DISPLAY "FOUND"
           END-SEARCH.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["FOUND"]);
}

#[test]
fn test_table_search_runtime_not_found() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-SRCH-NF.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS.
          05 ELEM OCCURS 3 TIMES INDEXED BY I PIC 9(2).
       PROCEDURE DIVISION.
           MOVE 11 TO ELEM(1).
           MOVE 22 TO ELEM(2).
           MOVE 33 TO ELEM(3).
           SEARCH WS
               AT END DISPLAY "NONE"
               WHEN ELEM(I) = 99 DISPLAY "FOUND"
           END-SEARCH.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["NONE"]);
}

#[test]
fn test_table_indexed_set_up_down_paths() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-SETPATH.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ARR.
          05 ELEM OCCURS 5 TIMES INDEXED BY IDX PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           SET IDX TO 1.
           MOVE 1 TO ELEM(IDX).
           SET IDX UP BY 2.
           MOVE 3 TO ELEM(IDX).
           SET IDX DOWN BY 1.
           DISPLAY ELEM(1).
           DISPLAY ELEM(2).
           DISPLAY ELEM(3).
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["1", "0", "3"]);
}

#[test]
fn test_table_copy_between_tables() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-COPY.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 SRC.
          05 SRC-ELEM OCCURS 3 TIMES PIC X(2).
       01 DST.
          05 DST-ELEM OCCURS 3 TIMES PIC X(2).
       PROCEDURE DIVISION.
           MOVE "AA" TO SRC-ELEM(1).
           MOVE "BB" TO SRC-ELEM(2).
           MOVE "CC" TO SRC-ELEM(3).
           MOVE SRC TO DST.
           DISPLAY DST-ELEM(1).
           DISPLAY DST-ELEM(2).
           DISPLAY DST-ELEM(3).
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["AA", "BB", "CC"]);
}
