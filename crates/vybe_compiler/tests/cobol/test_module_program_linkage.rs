use super::helpers::compile_ok;

#[test]
fn main_program_calls_worker_program_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. MAIN-PROG.\nPROCEDURE DIVISION.\n    CALL \"SUBPROG1\".\n    STOP RUN.",
    );
}

#[test]
fn call_using_multiple_args_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. MAIN-PROG-ARGS.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-NAME PIC X(10) VALUE \"ALICE\".\n01 WS-ID PIC 9(5) VALUE 42.\nPROCEDURE DIVISION.\n    CALL \"SUBPROG2\" USING WS-NAME WS-ID.\n    STOP RUN.",
    );
}

#[test]
fn call_with_on_exception_handling_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. MAIN-PROG-EX.\nPROCEDURE DIVISION.\n    CALL \"SUBFAIL\"\n        ON EXCEPTION DISPLAY \"FAIL\"\n        NOT ON EXCEPTION DISPLAY \"OK\"\n    END-CALL.\n    STOP RUN.",
    );
}

#[test]
fn nested_program_end_program_integration_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. OUTER-MODULE.\nPROCEDURE DIVISION.\n    DISPLAY \"OUTER\".\n    STOP RUN.\nEND PROGRAM OUTER-MODULE.",
    );
}

#[test]
fn call_using_single_argument_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. MAIN-ONE-ARG.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-V PIC 9(3) VALUE 7.\nPROCEDURE DIVISION.\n    CALL \"SUBPROG3\" USING WS-V.\n    STOP RUN.",
    );
}

#[test]
fn call_chain_with_two_program_names_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. MAIN-CHAIN.\nPROCEDURE DIVISION.\n    CALL \"SUB1\".\n    CALL \"SUB2\".\n    STOP RUN.",
    );
}

#[test]
fn module_call_with_using_and_exception_branches_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. MAIN-LINKAGE-EX.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-ID PIC 9(3) VALUE 101.\nPROCEDURE DIVISION.\n    CALL \"SUBPROG4\" USING WS-ID\n        ON EXCEPTION DISPLAY \"FAIL\"\n        NOT ON EXCEPTION DISPLAY \"OK\"\n    END-CALL.\n    STOP RUN.",
    );
}

#[test]
fn module_call_chain_with_three_steps_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. MAIN-CHAIN-3.\nPROCEDURE DIVISION.\n    CALL \"SUBA\".\n    CALL \"SUBB\".\n    CALL \"SUBC\".\n    STOP RUN.",
    );
}
