use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn national_usage_decl_compiles() {
    compile_ok(&p("01 N PIC X(30) USAGE NATIONAL.", "    DISPLAY N."));
}
#[test]
fn national_move_compiles() {
    compile_ok(&p(
        "01 A PIC X(20) USAGE NATIONAL VALUE \"TXT\".\n01 B PIC X(20) USAGE NATIONAL.",
        "    MOVE A TO B.",
    ));
}
#[test]
fn national_display_compiles() {
    compile_ok(&p(
        "01 N PIC X(20) USAGE NATIONAL VALUE \"HELLO\".",
        "    DISPLAY N.",
    ));
}
#[test]
fn special_names_alphabet_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    ALPHABET A1 IS STANDARD-1.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC X(5).\nPROCEDURE DIVISION.\n    DISPLAY X.\n    STOP RUN.",
    );
}
#[test]
fn special_names_decimal_point_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    DECIMAL-POINT IS COMMA.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC 9(3)V99.\nPROCEDURE DIVISION.\n    MOVE 1 TO X.\n    STOP RUN.",
    );
}
#[test]
fn xml_encoding_clause_compiles() {
    compile_ok(&p(
        "01 R PIC X(10).\n01 X PIC X(100).\n01 L PIC 9(5).",
        "    XML GENERATE X FROM R COUNT IN L ENCODING 1208.",
    ));
}
#[test]
fn xml_declaration_clause_compiles() {
    compile_ok(&p(
        "01 R PIC X(10).\n01 X PIC X(100).\n01 L PIC 9(5).",
        "    XML GENERATE X FROM R COUNT IN L WITH XML-DECLARATION.",
    ));
}
#[test]
fn xml_attributes_clause_compiles() {
    compile_ok(&p(
        "01 R.\n   05 A PIC X(5).\n01 X PIC X(100).\n01 L PIC 9(5).",
        "    XML GENERATE X FROM R COUNT IN L WITH ATTRIBUTES.",
    ));
}
#[test]
fn json_unicode_data_compiles() {
    compile_ok(&p(
        "01 J PIC X(100).\n01 R PIC X(20) USAGE NATIONAL.",
        "    JSON PARSE J INTO R.",
    ));
}
#[test]
fn trim_national_compiles() {
    compile_ok(&p(
        "01 N PIC X(20) USAGE NATIONAL.\n01 O PIC X(20) USAGE NATIONAL.",
        "    MOVE FUNCTION TRIM(N) TO O.",
    ));
}
#[test]
fn upper_national_compiles() {
    compile_ok(&p(
        "01 N PIC X(20) VALUE \"abc\".\n01 O PIC X(20).",
        "    MOVE FUNCTION UPPER-CASE(N) TO O.",
    ));
}
#[test]
fn lower_national_compiles() {
    compile_ok(&p(
        "01 N PIC X(20) VALUE \"ABC\".\n01 O PIC X(20).",
        "    MOVE FUNCTION LOWER-CASE(N) TO O.",
    ));
}
#[test]
fn length_national_compiles() {
    compile_ok(&p(
        "01 N PIC X(20) VALUE \"A\".\n01 L PIC 9(3).",
        "    MOVE FUNCTION LENGTH(N) TO L.",
    ));
}

#[test]
fn national_literal_move_compiles() {
    compile_ok(&p(
        "01 N1 PIC X(20) USAGE NATIONAL VALUE \"HELLO\".\n01 N2 PIC X(20) USAGE NATIONAL.",
        "    MOVE N1 TO N2.\n    DISPLAY N2.",
    ));
}

#[test]
fn national_to_display_move_compiles() {
    compile_ok(&p(
        "01 N PIC X(20) USAGE NATIONAL VALUE \"HELLO\".\n01 D PIC X(20).",
        "    MOVE N TO D.\n    DISPLAY D.",
    ));
}
