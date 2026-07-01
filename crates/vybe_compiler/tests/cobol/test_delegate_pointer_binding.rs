use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn function_pointer_declaration_compiles() {
    compile_ok(&p(
        "01 WS-FPTR USAGE IS FUNCTION-POINTER.",
        "    DISPLAY \"FPTR\".",
    ));
}

#[test]
fn procedure_pointer_declaration_compiles() {
    compile_ok(&p(
        "01 WS-PPTR USAGE IS PROCEDURE-POINTER.",
        "    DISPLAY \"PPTR\".",
    ));
}

#[test]
fn generic_pointer_and_delegate_slot_compile() {
    compile_ok(&p(
        "01 WS-PTR USAGE IS POINTER.\n01 WS-CALLBACK USAGE IS PROCEDURE-POINTER.",
        "    SET WS-PTR TO NULL.\n    DISPLAY \"READY\".",
    ));
}

#[test]
fn callback_style_call_with_pointer_args_compiles() {
    compile_ok(&p(
        "01 WS-CALLBACK USAGE IS PROCEDURE-POINTER.\n01 WS-ARG PIC X(10) VALUE \"PAYLOAD\".",
        "    CALL \"INVOKE-CALLBACK\" USING WS-CALLBACK WS-ARG.",
    ));
}

#[test]
fn set_procedure_pointer_to_entry_compiles() {
    compile_ok(&p(
        "01 P USAGE IS PROCEDURE-POINTER.",
        "    SET P TO ENTRY \"WORKER\".",
    ));
}

#[test]
fn call_through_procedure_pointer_compiles() {
    compile_ok(&p(
        "01 P USAGE IS PROCEDURE-POINTER.",
        "    CALL P.",
    ));
}

#[test]
fn delegate_swap_two_pointers_compiles() {
    compile_ok(&p(
        "01 P1 USAGE IS PROCEDURE-POINTER.\n01 P2 USAGE IS PROCEDURE-POINTER.",
        "    SET P1 TO ENTRY \"A\".\n    SET P2 TO ENTRY \"B\".\n    MOVE P1 TO P2.",
    ));
}

#[test]
fn delegate_call_with_payload_compiles() {
    compile_ok(&p(
        "01 P USAGE IS PROCEDURE-POINTER.\n01 PAY PIC X(10) VALUE \"DATA\".",
        "    CALL P USING PAY.",
    ));
}
