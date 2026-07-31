use crate::helpers;

#[test]
fn test_search_linear_basic() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-LINEAR.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 FILLER PIC X(3) VALUE "AAA".
          05 FILLER PIC X(3) VALUE "BBB".
          05 FILLER PIC X(3) VALUE "CCC".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 3 TIMES INDEXED BY IDX.
             10 VAL PIC X(3).
       PROCEDURE DIVISION.
           SET IDX TO 1.
           SEARCH TBL-ENTRY
              AT END DISPLAY "NOT FOUND"
              WHEN VAL(IDX) = "BBB"
                 DISPLAY "FOUND " VAL(IDX)
           END-SEARCH.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["FOUND BBB"]);
}

#[test]
fn test_search_linear_not_found() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-NOT-FOUND.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 FILLER PIC X(3) VALUE "AAA".
          05 FILLER PIC X(3) VALUE "BBB".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 2 TIMES INDEXED BY IDX.
             10 VAL PIC X(3).
       PROCEDURE DIVISION.
           SET IDX TO 1.
           SEARCH TBL-ENTRY
              AT END DISPLAY "NOT FOUND"
              WHEN VAL(IDX) = "CCC"
                 DISPLAY "FOUND"
           END-SEARCH.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["NOT FOUND"]);
}

#[test]
fn test_search_all_binary() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-ALL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 FILLER PIC X(5) VALUE "01AAA".
          05 FILLER PIC X(5) VALUE "02BBB".
          05 FILLER PIC X(5) VALUE "03CCC".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 3 TIMES
             ASCENDING KEY IS KEY-ID
             INDEXED BY IDX.
             10 KEY-ID PIC 9(2).
             10 VAL PIC X(3).
       PROCEDURE DIVISION.
           SEARCH ALL TBL-ENTRY
              AT END DISPLAY "NOT FOUND"
              WHEN KEY-ID(IDX) = 03
                 DISPLAY VAL(IDX)
           END-SEARCH.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["CCC"]);
}

#[test]
fn test_search_all_descending() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-ALL-DESC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 FILLER PIC X(5) VALUE "03CCC".
          05 FILLER PIC X(5) VALUE "02BBB".
          05 FILLER PIC X(5) VALUE "01AAA".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 3 TIMES
             DESCENDING KEY IS KEY-ID
             INDEXED BY IDX.
             10 KEY-ID PIC 9(2).
             10 VAL PIC X(3).
       PROCEDURE DIVISION.
           SEARCH ALL TBL-ENTRY
              AT END DISPLAY "NOT FOUND"
              WHEN KEY-ID(IDX) = 01
                 DISPLAY VAL(IDX)
           END-SEARCH.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["AAA"]);
}

#[test]
fn test_search_linear_varying() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-VARYING.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 FILLER PIC X VALUE "X".
          05 FILLER PIC X VALUE "Y".
          05 FILLER PIC X VALUE "Z".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 3 TIMES INDEXED BY IDX.
             10 VAL PIC X.
       01 TRACKER PIC 9 VALUE 1.
       PROCEDURE DIVISION.
           SET IDX TO 1.
           SEARCH TBL-ENTRY VARYING TRACKER
              AT END DISPLAY "NOT FOUND"
              WHEN VAL(IDX) = "Z"
                 DISPLAY TRACKER
           END-SEARCH.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["3"]);
}

#[test]
fn test_search_multiple_when() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-MULTI-WHEN.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 FILLER PIC X VALUE "A".
          05 FILLER PIC X VALUE "B".
          05 FILLER PIC X VALUE "C".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 3 TIMES INDEXED BY IDX.
             10 VAL PIC X.
       PROCEDURE DIVISION.
           SET IDX TO 1.
           SEARCH TBL-ENTRY
              AT END DISPLAY "NOT FOUND"
              WHEN VAL(IDX) = "X" DISPLAY "FOUND X"
              WHEN VAL(IDX) = "B" DISPLAY "FOUND B"
           END-SEARCH.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["FOUND B"]);
}

#[test]
fn test_search_linear_with_if_no_match() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-IF-NOMATCH.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 FILLER PIC X VALUE "1".
          05 FILLER PIC X VALUE "2".
          05 FILLER PIC X VALUE "3".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 3 TIMES INDEXED BY IDX.
             10 VAL PIC X.
       01 RESULT PIC X VALUE "N".
       PROCEDURE DIVISION.
           SET IDX TO 1.
           SEARCH TBL-ENTRY
              AT END IF RESULT = "N" DISPLAY "NO" END-IF
              WHEN VAL(IDX) = "9"
                 MOVE "Y" TO RESULT
                 DISPLAY RESULT
           END-SEARCH.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["NO"]);
}

#[test]
fn test_search_all_uses_ascending_key_order() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-ALL-ORDER.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 FILLER PIC X(5) VALUE "05PEND".
          05 FILLER PIC X(5) VALUE "02CODE".
          05 FILLER PIC X(5) VALUE "07END ".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 3 TIMES
             ASCENDING KEY IS KEY-ID
             INDEXED BY IDX.
             10 KEY-ID PIC 9(2).
             10 VAL PIC X(3).
       PROCEDURE DIVISION.
           SEARCH ALL TBL-ENTRY
              WHEN KEY-ID(IDX) = 07
                 DISPLAY VAL(IDX)
           END-SEARCH.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["END"]);
}

#[test]
fn test_search_all_duplicate_keys_prefers_first_match() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-ALL-DUP.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 FILLER PIC X(6) VALUE "01A".
          05 FILLER PIC X(6) VALUE "01B".
          05 FILLER PIC X(6) VALUE "02C".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 3 TIMES
             ASCENDING KEY IS KEY-ID
             INDEXED BY IDX.
             10 KEY-ID PIC 9(2).
             10 VAL PIC X(4).
       PROCEDURE DIVISION.
           SEARCH ALL TBL-ENTRY
              WHEN KEY-ID(IDX) = 01
                 DISPLAY VAL(IDX)
           END-SEARCH.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["A"]);
}
