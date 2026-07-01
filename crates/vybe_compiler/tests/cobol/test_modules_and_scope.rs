use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn nested_program_scope_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. PARENT.\nPROCEDURE DIVISION.\n    DISPLAY \"PARENT\".\n    STOP RUN.\nEND PROGRAM PARENT.\nIDENTIFICATION DIVISION.\nPROGRAM-ID. CHILD.\nPROCEDURE DIVISION.\n    DISPLAY \"CHILD\".\n    STOP RUN.\nEND PROGRAM CHILD.",
    );
}

#[test]
fn separate_program_sections_compile() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-X PIC 9(1) VALUE 1.\nPROCEDURE DIVISION.\n    DISPLAY WS-X.\n    STOP RUN.",
    );
}

#[test]
fn paragraph_scope_with_perform_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM PARA-ONE.\n    STOP RUN.\nPARA-ONE.\n    DISPLAY \"ONE\".",
    );
}

#[test]
fn section_scope_with_perform_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM SECTION-ONE.\n    STOP RUN.\nSECTION-ONE SECTION.\n    DISPLAY \"ONE\".",
    );
}
