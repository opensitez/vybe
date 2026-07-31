use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn json_generate_feature_compiles() {
    compile_ok(&p(
        "01 WS-REC.\n   05 WS-NAME PIC X(10) VALUE \"ALICE\".\n   05 WS-AGE PIC 9(3) VALUE 30.\n01 WS-JSON PIC X(200).",
        "    JSON GENERATE WS-JSON FROM WS-REC.",
    ));
}

#[test]
fn json_parse_feature_compiles() {
    compile_ok(&p(
        "01 WS-JSON PIC X(200) VALUE '{\"name\":\"BOB\"}'.\n01 WS-REC.\n   05 WS-NAME PIC X(10).",
        "    JSON PARSE WS-JSON INTO WS-REC.",
    ));
}

#[test]
fn xml_generate_feature_compiles() {
    compile_ok(&p(
        "01 WS-REC.\n   05 WS-NAME PIC X(10) VALUE \"ALICE\".\n01 WS-XML PIC X(500).\n01 WS-LEN PIC 9(5).",
        "    XML GENERATE WS-XML FROM WS-REC COUNT IN WS-LEN.",
    ));
}

#[test]
fn xml_parse_feature_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-XML PIC X(200) VALUE \"<a>1</a>\".\nPROCEDURE DIVISION.\n    XML PARSE WS-XML PROCESSING PROCEDURE X-HANDLER.\n    STOP RUN.\nX-HANDLER SECTION.\n    DISPLAY \"TAG\".",
    );
}

#[test]
fn embedded_sql_select_compiles() {
    compile_ok(&p(
        "01 WS-ID PIC 9(10) VALUE 1.\n01 WS-NAME PIC X(50).",
        "    EXEC SQL SELECT NAME INTO :WS-NAME FROM USERS WHERE ID = :WS-ID END-EXEC.",
    ));
}

#[test]
fn embedded_sql_insert_compiles() {
    compile_ok(&p(
        "01 WS-ID PIC 9(10) VALUE 1.\n01 WS-NAME PIC X(50) VALUE \"A\".",
        "    EXEC SQL INSERT INTO USERS (ID, NAME) VALUES (:WS-ID, :WS-NAME) END-EXEC.",
    ));
}

#[test]
fn json_generate_with_nested_group_compiles() {
    compile_ok(&p(
        "01 WS-REC.\n   05 WS-NAME PIC X(10) VALUE \"NINA\".\n   05 WS-ADDR.\n      10 WS-CITY PIC X(10) VALUE \"PARIS\".\n01 WS-JSON PIC X(400).",
        "    JSON GENERATE WS-JSON FROM WS-REC.",
    ));
}

#[test]
fn xml_generate_with_declaration_compiles() {
    compile_ok(&p(
        "01 WS-REC.\n   05 WS-ID PIC 9(4) VALUE 1001.\n01 WS-XML PIC X(500).\n01 WS-LEN PIC 9(5).",
        "    XML GENERATE WS-XML FROM WS-REC COUNT IN WS-LEN WITH XML-DECLARATION.",
    ));
}

#[test]
fn json_generate_with_count_compiles() {
    compile_ok(&p(
        "01 WS-REC.\n   05 WS-NAME PIC X(8) VALUE \"BOB\".\n01 WS-JSON PIC X(200).\n01 WS-LEN PIC 9(5).",
        "    JSON GENERATE WS-JSON FROM WS-REC COUNT IN WS-LEN.",
    ));
}

#[test]
fn xml_generate_with_attributes_compiles() {
    compile_ok(&p(
        "01 WS-REC.\n   05 WS-ID PIC 9(4) VALUE 1001.\n01 WS-XML PIC X(500).",
        "    XML GENERATE WS-XML FROM WS-REC WITH ATTRIBUTES.",
    ));
}
