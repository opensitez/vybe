use crate::helpers;

#[test]
fn test_json_generate_basic() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. JSON-GEN.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 REC.
          05 FLD-A PIC X(5) VALUE "HELLO".
          05 FLD-B PIC 9(3) VALUE 123.
       01 JSON-DOC PIC X(50).
       PROCEDURE DIVISION.
           JSON GENERATE JSON-DOC FROM REC.
           DISPLAY "JSON GEN PARSED".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["JSON GEN PARSED"]);
}

#[test]
fn test_json_generate_exception() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. JSON-GEN-EXC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 REC PIC X(5) VALUE "HELLO".
       01 JSON-DOC PIC X(2).
       PROCEDURE DIVISION.
           JSON GENERATE JSON-DOC FROM REC
              ON EXCEPTION DISPLAY "EXCEPTION CAUGHT"
              NOT ON EXCEPTION DISPLAY "SUCCESS".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["EXCEPTION CAUGHT"]);
}

#[test]
fn test_json_parse_basic() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. JSON-PARSE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 JSON-DOC PIC X(20) VALUE '{"A":"HI"}'.
       01 REC.
          05 A PIC X(2).
       PROCEDURE DIVISION.
           JSON PARSE JSON-DOC INTO REC.
           DISPLAY A.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["HI"]);
}

#[test]
fn test_json_parse_with_detail() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. JSON-PARSE-DETAIL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 JSON-DOC PIC X(20) VALUE '{"A":"HI"}'.
       01 REC.
          05 A PIC X(2).
       PROCEDURE DIVISION.
           JSON PARSE JSON-DOC INTO REC
              WITH DETAIL
              ON EXCEPTION DISPLAY "EXC".
           DISPLAY A.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["HI"]);
}

#[test]
fn test_xml_generate_basic() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. XML-GEN.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 REC.
          05 FLD PIC X(5) VALUE "HELLO".
       01 XML-DOC PIC X(50).
       PROCEDURE DIVISION.
           XML GENERATE XML-DOC FROM REC.
           DISPLAY "XML GEN PARSED".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["XML GEN PARSED"]);
}

#[test]
fn test_xml_parse_basic() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. XML-PARSE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 XML-DOC PIC X(50) VALUE "<REC><FLD>HI</FLD></REC>".
       PROCEDURE DIVISION.
           XML PARSE XML-DOC
              PROCESSING PROCEDURE XML-PROC.
           DISPLAY "XML PARSED".
           STOP RUN.
       XML-PROC SECTION.
           EXIT.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["XML PARSED"]);
}
