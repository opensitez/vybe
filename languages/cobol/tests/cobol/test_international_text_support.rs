use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn national_usage_data_item_compiles() {
    compile_ok(&p(
        "01 WS-TEXT PIC X(50) USAGE NATIONAL.",
        "    DISPLAY WS-TEXT.",
    ));
}

#[test]
fn national_text_move_and_display_compiles() {
    compile_ok(&p(
        "01 WS-SRC PIC X(20) USAGE NATIONAL VALUE \"Unicode\".\n01 WS-DST PIC X(20) USAGE NATIONAL.",
        "    MOVE WS-SRC TO WS-DST.\n    DISPLAY WS-DST.",
    ));
}

#[test]
fn xml_generate_with_encoding_clause_compiles() {
    compile_ok(&p(
        "01 WS-REC.\n   05 WS-NAME PIC X(10) VALUE \"ALICE\".\n01 WS-XML PIC X(200).\n01 WS-LEN PIC 9(5).",
        "    XML GENERATE WS-XML FROM WS-REC COUNT IN WS-LEN ENCODING 1208.",
    ));
}

#[test]
fn special_names_with_alphabet_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    ALPHABET MY-ALPHA IS STANDARD-1.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-TXT PIC X(10) VALUE \"ABC\".\nPROCEDURE DIVISION.\n    DISPLAY WS-TXT.\n    STOP RUN.",
    );
}
