use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn pointer_variable_declares() {
    compile_ok(&p(
        "01 WS-PTR USAGE POINTER.\n01 WS-VAL PIC X(5) VALUE \"DATA\".",
        "    SET WS-PTR TO NULL.",
    ));
}

#[test]
fn pointer_to_data_compiles() {
    compile_ok(&p(
        "01 WS-PTR USAGE POINTER.\n01 WS-VAL PIC X(5) VALUE \"DATA\".",
        "    SET WS-PTR TO ADDRESS OF WS-VAL.",
    ));
}

#[test]
fn pointer_move_between_pointer_variables_compiles() {
    compile_ok(&p(
        "01 P1 USAGE POINTER.\n01 P2 USAGE POINTER.\n01 BUF PIC X(4) VALUE \"DATA\".",
        "    SET P1 TO ADDRESS OF BUF.\n    MOVE P1 TO P2.",
    ));
}

#[test]
fn pointer_null_comparison_in_if_compiles() {
    compile_ok(&p(
        "01 P USAGE POINTER.",
        "    SET P TO NULL.\n    IF P = NULL\n        DISPLAY \"NULL\"\n    END-IF.",
    ));
}
