use super::helpers::compile_ok;

// ── XML GENERATE basic ────────────────────────────────────────

#[test]
fn xml_generate_basic() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-record.
           05 ws-name  PIC X(20) VALUE "Alice".
           05 ws-age   PIC 99    VALUE 30.
       01 ws-xml   PIC X(500).
       01 ws-count PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-record
               COUNT IN ws-count
           DISPLAY ws-xml(1:ws-count)
           STOP RUN.
"#,
    );
}

#[test]
fn xml_generate_single_field() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-title PIC X(30) VALUE "COBOL XML Test".
       01 ws-xml   PIC X(200).
       01 ws-len   PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-title
               COUNT IN ws-len
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.
"#,
    );
}

#[test]
fn xml_generate_numeric_field() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-amount PIC 9(7)V99 VALUE 12345.67.
       01 ws-xml    PIC X(200).
       01 ws-len    PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-amount
               COUNT IN ws-len
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.
"#,
    );
}

#[test]
fn xml_generate_nested_group() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-person.
           05 ws-first-name  PIC X(15) VALUE "John".
           05 ws-last-name   PIC X(20) VALUE "Smith".
           05 ws-address.
               10 ws-street  PIC X(30) VALUE "123 Main St".
               10 ws-city    PIC X(20) VALUE "Springfield".
               10 ws-state   PIC XX    VALUE "IL".
               10 ws-zip     PIC X(10) VALUE "62701".
       01 ws-xml  PIC X(1000).
       01 ws-len  PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-person
               COUNT IN ws-len
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.
"#,
    );
}

#[test]
fn xml_generate_with_encoding() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-data.
           05 ws-id    PIC 9(5) VALUE 42.
           05 ws-label PIC X(10) VALUE "item".
       01 ws-xml  PIC X(500).
       01 ws-len  PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-data
               COUNT IN ws-len
               ENCODING 1208
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.
"#,
    );
}

#[test]
fn xml_generate_with_xml_declaration() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-item.
           05 ws-name PIC X(10) VALUE "widget".
           05 ws-qty  PIC 99    VALUE 5.
       01 ws-xml PIC X(500).
       01 ws-len PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-item
               COUNT IN ws-len
               WITH XML-DECLARATION
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.
"#,
    );
}

#[test]
fn xml_generate_with_attributes() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-product.
           05 ws-id    PIC 9(5)  VALUE 1001.
           05 ws-name  PIC X(20) VALUE "Widget".
           05 ws-price PIC 9(5)V99 VALUE 9.99.
       01 ws-xml PIC X(500).
       01 ws-len PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-product
               COUNT IN ws-len
               WITH ATTRIBUTES
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.
"#,
    );
}

#[test]
fn xml_generate_on_exception() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-rec.
           05 ws-val PIC X(5) VALUE "test".
       01 ws-xml    PIC X(50).
       01 ws-len    PIC 9(5).
       01 ws-err    PIC X VALUE "N".
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-rec
               COUNT IN ws-len
               ON EXCEPTION     MOVE "Y" TO ws-err
               NOT ON EXCEPTION MOVE "N" TO ws-err
           END-XML
           DISPLAY ws-err
           STOP RUN.
"#,
    );
}

// ── XML PARSE ────────────────────────────────────────────────

#[test]
fn xml_parse_basic() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-xml-doc PIC X(200)
           VALUE "<person><name>Alice</name><age>30</age></person>".
       01 ws-event-count PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           XML PARSE ws-xml-doc
               PROCESSING PROCEDURE xml-handler
           DISPLAY ws-event-count
           STOP RUN.
       xml-handler SECTION.
           ADD 1 TO ws-event-count.
"#,
    );
}

#[test]
fn xml_parse_with_encoding() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-xml PIC X(100)
           VALUE "<root><item>value</item></root>".
       01 ws-count PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           XML PARSE ws-xml
               PROCESSING PROCEDURE handle-xml
               ENCODING 1208
           DISPLAY ws-count
           STOP RUN.
       handle-xml SECTION.
           ADD 1 TO ws-count.
"#,
    );
}

#[test]
fn xml_parse_on_exception() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-xml  PIC X(50) VALUE "<valid>data</valid>".
       01 ws-err  PIC X     VALUE "N".
       01 ws-ok   PIC X     VALUE "N".
       PROCEDURE DIVISION.
           XML PARSE ws-xml
               PROCESSING PROCEDURE xml-proc
               ON EXCEPTION     MOVE "Y" TO ws-err
               NOT ON EXCEPTION MOVE "Y" TO ws-ok
           END-XML
           DISPLAY ws-ok
           STOP RUN.
       xml-proc SECTION.
           CONTINUE.
"#,
    );
}

#[test]
fn xml_parse_event_types() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-xml PIC X(200)
           VALUE "<book><title>COBOL Guide</title><pages>400</pages></book>".
       01 ws-start-count   PIC 99 VALUE 0.
       01 ws-end-count     PIC 99 VALUE 0.
       01 ws-content-count PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           XML PARSE ws-xml
               PROCESSING PROCEDURE parse-events
           DISPLAY ws-start-count
           DISPLAY ws-end-count
           DISPLAY ws-content-count
           STOP RUN.
       parse-events SECTION.
           EVALUATE XML-CODE
               WHEN "START-OF-ELEMENT"
                   ADD 1 TO ws-start-count
               WHEN "END-OF-ELEMENT"
                   ADD 1 TO ws-end-count
               WHEN "CONTENT-CHARACTERS"
                   ADD 1 TO ws-content-count
               WHEN OTHER
                   CONTINUE
           END-EVALUATE.
"#,
    );
}

#[test]
fn xml_parse_extract_content() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-xml PIC X(100)
           VALUE "<employee><name>Bob</name><dept>IT</dept></employee>".
       01 ws-current-element PIC X(20).
       01 ws-name-val        PIC X(20).
       01 ws-dept-val        PIC X(20).
       PROCEDURE DIVISION.
           XML PARSE ws-xml
               PROCESSING PROCEDURE extract-data
           DISPLAY ws-name-val
           DISPLAY ws-dept-val
           STOP RUN.
       extract-data SECTION.
           EVALUATE XML-CODE
               WHEN "START-OF-ELEMENT"
                   MOVE XML-TEXT TO ws-current-element
               WHEN "CONTENT-CHARACTERS"
                   EVALUATE ws-current-element
                       WHEN "name" MOVE XML-TEXT TO ws-name-val
                       WHEN "dept" MOVE XML-TEXT TO ws-dept-val
                   END-EVALUATE
               WHEN OTHER CONTINUE
           END-EVALUATE.
"#,
    );
}

// ── XML GENERATE + PARSE roundtrip ───────────────────────────

#[test]
fn xml_generate_parse_roundtrip() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-original.
           05 ws-id   PIC 9(5) VALUE 42.
           05 ws-name PIC X(10) VALUE "Alice".
       01 ws-xml      PIC X(500).
       01 ws-xml-len  PIC 9(5).
       01 ws-events   PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-original
               COUNT IN ws-xml-len
           XML PARSE ws-xml(1:ws-xml-len)
               PROCESSING PROCEDURE count-events
           DISPLAY ws-events
           STOP RUN.
       count-events SECTION.
           ADD 1 TO ws-events.
"#,
    );
}

#[test]
fn xml_generate_special_chars() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-data.
           05 ws-note PIC X(30) VALUE "Price < 100 & > 50".
       01 ws-xml PIC X(500).
       01 ws-len PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-data
               COUNT IN ws-len
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.
"#,
    );
}

#[test]
fn xml_generate_name_override() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-cust-record.
           05 cust-id    PIC 9(5) VALUE 101.
           05 cust-name  PIC X(20) VALUE "John Doe".
           05 cust-email PIC X(30) VALUE "john@example.com".
       01 ws-xml PIC X(500).
       01 ws-len PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-cust-record
               COUNT IN ws-len
               NAMESPACE IS "http://example.com/cust"
               NAMESPACE-PREFIX IS "cust"
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.
"#,
    );
}

#[test]
fn xml_generate_suppress_when_zero() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-order.
           05 ws-order-id  PIC 9(5) VALUE 500.
           05 ws-qty       PIC 99   VALUE 0.
           05 ws-amount    PIC 9(7)V99 VALUE 0.
       01 ws-xml PIC X(500).
       01 ws-len PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-order
               COUNT IN ws-len
               SUPPRESS WHEN ZERO
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.
"#,
    );
}

#[test]
fn xml_generate_suppress_when_spaces() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-contact.
           05 ws-name  PIC X(20) VALUE "Alice".
           05 ws-phone PIC X(15) VALUE SPACES.
           05 ws-email PIC X(30) VALUE SPACES.
       01 ws-xml PIC X(500).
       01 ws-len PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-contact
               COUNT IN ws-len
               SUPPRESS WHEN SPACES
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.
"#,
    );
}
