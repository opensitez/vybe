use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn raise_exception_compiles() {
    compile_ok(&p("", "    RAISE EXCEPTION \"boom\"."));
}

#[test]
fn raise_and_display_compiles() {
    compile_ok(&p(
        "",
        "    RAISE EXCEPTION \"fatal\".\n    DISPLAY \"after error\".",
    ));
}

#[test]
fn call_with_exception_path_compiles() {
    compile_ok(&p("", "    CALL \"SUB\".\n    DISPLAY \"done\"."));
}

#[test]
fn call_end_call_exception_branches_compile() {
    compile_ok(&p(
        "",
        "    CALL \"WORK\"\n        ON EXCEPTION DISPLAY \"ERR\"\n        NOT ON EXCEPTION DISPLAY \"OK\"\n    END-CALL.",
    ));
}

#[test]
fn raise_exception_followed_by_branch_compile() {
    compile_ok(&p(
        "01 WS-ERR PIC X(4) VALUE \"FAIL\".",
        "    IF WS-ERR = \"FAIL\"\n        RAISE EXCEPTION \"boom\"\n    END-IF.",
    ));
}

#[test]
fn call_with_exception_and_recovery_path_compiles() {
    compile_ok(&p(
        "01 WS-STATUS PIC X(2) VALUE \"OK\".",
        "    CALL \"MAY-FAIL\"\n        ON EXCEPTION MOVE \"ER\" TO WS-STATUS\n    END-CALL.\n    IF WS-STATUS = \"ER\" DISPLAY \"RECOVER\" END-IF.",
    ));
}

#[test]
fn raise_exception_in_conditional_flow_compiles() {
    compile_ok(&p(
        "01 FLAG PIC 9 VALUE 1.",
        "    IF FLAG = 1\n        RAISE EXCEPTION \"E1\"\n    ELSE\n        DISPLAY \"NOERR\"\n    END-IF.",
    ));
}
