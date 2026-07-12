use crate::helpers;

#[test]
fn test_xml_parse_exception_handler() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. XML-EXC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 XML-DOC PIC X(50) VALUE "<BAD>XML".
       PROCEDURE DIVISION.
           XML PARSE XML-DOC
              PROCESSING PROCEDURE XML-PROC
              ON EXCEPTION DISPLAY "EXCEPTION CAUGHT"
              NOT ON EXCEPTION DISPLAY "SUCCESS".
           STOP RUN.
       XML-PROC SECTION.
           EXIT.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_json_parse_name_suppress() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. JSON-NAME.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 JSON-DOC PIC X(50) VALUE '{"A":"HI"}'.
       01 REC.
          05 FLD-A PIC X(2) NAME IS "A".
       PROCEDURE DIVISION.
           JSON PARSE JSON-DOC INTO REC
              SUPPRESS FLD-A.
           DISPLAY "JSON SUPPRESS PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}
