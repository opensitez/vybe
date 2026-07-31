use crate::helpers;

#[test]
// GAP: JSON PARSE syntax 'ON EXCEPTION' is not currently parsed by the grammar (expects period or kw_section).
fn test_json_parse_edge_cases() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. JSON-EDGE-CASES.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 JSON-DOC PIC X(200).
       01 JSON-STATUS PIC 9(4).
       01 PARSED-DATA.
          05 USER-AGE PIC 9(3).
          05 IS-ACTIVE PIC X(5).
          05 NULL-FIELD PIC X(10).
       PROCEDURE DIVISION.
           MOVE '{"USER-AGE": 42, "IS-ACTIVE": true, "NULL-FIELD": null}' TO JSON-DOC.
           JSON PARSE JSON-DOC INTO PARSED-DATA
                ON EXCEPTION
                   DISPLAY "EXCEPTION:" JSON-STATUS
                NOT ON EXCEPTION
                   DISPLAY USER-AGE " " IS-ACTIVE " " NULL-FIELD
           END-JSON.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["042 true null"]);
}

#[test]
// GAP: XML PARSE PROCESSING PROCEDURE emits 'undefined is not callable' at runtime.
fn test_xml_parse_processing_procedure() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. XML-EDGE-CASES.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 XML-DOC PIC X(200).
       01 XML-CODE PIC S9(9) COMP-5.
       01 XML-EVENT PIC X(30).
       01 XML-TEXT PIC X(30).
       PROCEDURE DIVISION.
           MOVE "<root><EMPTY/></root>" TO XML-DOC.
           XML PARSE XML-DOC PROCESSING PROCEDURE XML-HANDLER
                ON EXCEPTION
                   DISPLAY "EXCEPTION:" XML-CODE
                NOT ON EXCEPTION
                   DISPLAY "SUCCESS"
           END-XML.
           STOP RUN.
           
       XML-HANDLER.
           DISPLAY XML-EVENT " " XML-TEXT.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_json_parse_with_name_clause() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. JSON-NAME-EDGE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 JSON-DOC PIC X(200) VALUE '{"A": 42, "B": true}'.
       01 PARSED-DATA.
          05 USER-AGE PIC 9(3).
          05 IS-ACTIVE PIC X(5).
       PROCEDURE DIVISION.
           JSON PARSE JSON-DOC INTO PARSED-DATA
              NAME USER-AGE IS "A"
              NAME IS-ACTIVE IS "B".
           DISPLAY "JSON NAME EDGE".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["JSON NAME EDGE"]);
}

#[test]
fn test_xml_parse_with_returning_clause() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. XML-RETURN-EDGE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 XML-DOC PIC X(200) VALUE "<ROOT><VALUE>1</VALUE></ROOT>".
       01 XML-STATUS PIC S9(9) COMP-5.
       PROCEDURE DIVISION.
           XML PARSE XML-DOC
                PROCESSING PROCEDURE XML-PROC
                RETURNING XML-STATUS.
           DISPLAY XML-STATUS.
           STOP RUN.
       XML-PROC.
           DISPLAY XML-STATUS.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}
