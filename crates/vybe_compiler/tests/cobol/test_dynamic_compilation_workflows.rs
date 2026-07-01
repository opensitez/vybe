use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn dynamic_compile_external_service_call_compiles() {
    compile_ok(&p(
        "01 WS-SOURCE PIC X(200).\n01 WS-HANDLE PIC X(40).",
        "    CALL \"DYNAMIC-COMPILE\" USING WS-SOURCE RETURNING WS-HANDLE.",
    ));
}

#[test]
fn dynamic_load_and_invoke_pattern_compiles() {
    compile_ok(&p(
        "01 WS-MODULE PIC X(40) VALUE \"PLUGIN-A\".",
        "    CALL \"LOAD-MODULE\" USING WS-MODULE.\n    CALL \"INVOKE-MODULE\" USING WS-MODULE.",
    ));
}

#[test]
fn dynamic_rununit_switch_pattern_compiles() {
    compile_ok(&p(
        "01 WS-UNIT PIC X(20) VALUE \"U1\".",
        "    CALL \"SET-RUNUNIT\" USING WS-UNIT.\n    DISPLAY \"RUNUNIT-SET\".",
    ));
}

#[test]
fn dynamic_dispatch_with_exception_branch_compiles() {
    compile_ok(&p(
        "01 WS-TARGET PIC X(20) VALUE \"HANDLER\".",
        "    CALL WS-TARGET\n        ON EXCEPTION DISPLAY \"CALL-FAIL\"\n        NOT ON EXCEPTION DISPLAY \"CALL-OK\"\n    END-CALL.",
    ));
}
