use super::helpers::compile_ok;

fn program(body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n{}\n",
        body
    )
}

#[test]
fn declaratives_standard_error_handler_on_file() {
    compile_ok(&program(
        "DECLARATIVES.\nERR-SEC SECTION.\n    USE AFTER STANDARD ERROR PROCEDURE ON WS-FILE.\nEND DECLARATIVES.\nMAIN-SEC SECTION.\n    DISPLAY \"RUN\".\n    STOP RUN.",
    ));
}

#[test]
fn declaratives_debugging_on_all_procedures() {
    compile_ok(&program(
        "DECLARATIVES.\nDBG-SEC SECTION.\n    USE FOR DEBUGGING ON ALL PROCEDURES.\nEND DECLARATIVES.\nMAIN-SEC SECTION.\n    DISPLAY \"RUN\".\n    STOP RUN.",
    ));
}

#[test]
fn declaratives_with_multiple_sections() {
    compile_ok(&program(
        "DECLARATIVES.\nD-A SECTION.\n    USE AFTER STANDARD ERROR PROCEDURE ON WS-FILE.\nD-B SECTION.\n    USE AFTER STANDARD ERROR PROCEDURE ON WS-FILE.\nEND DECLARATIVES.\nMAIN SECTION.\n    DISPLAY \"OK\".\n    STOP RUN.",
    ));
}

#[test]
fn declaratives_empty_block_then_main_section() {
    compile_ok(&program(
        "DECLARATIVES.\nEND DECLARATIVES.\nMAIN SECTION.\n    DISPLAY \"OK\".\n    STOP RUN.",
    ));
}

#[test]
fn declaratives_flow_with_labeled_main_section() {
    compile_ok(&program(
        "DECLARATIVES.\nD1 SECTION.\n    USE AFTER STANDARD ERROR PROCEDURE ON WS-FILE.\nEND DECLARATIVES.\nM1 SECTION.\n    DISPLAY \"MAIN\".\n    STOP RUN.",
    ));
}

#[test]
fn declaratives_flow_with_call_in_main_section() {
    compile_ok(&program(
        "DECLARATIVES.\nD2 SECTION.\n    USE AFTER STANDARD ERROR PROCEDURE ON WS-FILE.\nEND DECLARATIVES.\nM2 SECTION.\n    CALL \"WORK\".\n    STOP RUN.",
    ));
}

#[test]
fn declaratives_debugging_does_not_break_simple_if() {
    compile_ok(&program(
        "DECLARATIVES.\nDBG-SEC SECTION.\n    USE FOR DEBUGGING ON ALL PROCEDURES.\nEND DECLARATIVES.\nMAIN SECTION.\n    IF 1 = 1 DISPLAY \"OK\" END-IF.\n    STOP RUN.",
    ));
}

#[test]
fn declaratives_standard_error_then_perform_loop() {
    compile_ok(&program(
        "DECLARATIVES.\nERR-SEC SECTION.\n    USE AFTER STANDARD ERROR PROCEDURE ON WS-FILE.\nEND DECLARATIVES.\nMAIN SECTION.\n    PERFORM 2 TIMES\n        DISPLAY \"X\"\n    END-PERFORM.\n    STOP RUN.",
    ));
}

#[test]
fn declaratives_standard_error_with_evaluate() {
    compile_ok(&program(
        "DECLARATIVES.\nERR-SEC SECTION.\n    USE AFTER STANDARD ERROR PROCEDURE ON WS-FILE.\nEND DECLARATIVES.\nMAIN SECTION.\n    EVALUATE TRUE\n        WHEN 1 = 1 DISPLAY \"A\"\n        WHEN OTHER DISPLAY \"B\"\n    END-EVALUATE.\n    STOP RUN.",
    ));
}
