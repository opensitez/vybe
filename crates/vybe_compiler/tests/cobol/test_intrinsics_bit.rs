use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_intrinsics_hex_conversions() {
    compile_ok(&p(
        r#"
01 WS-HEX PIC X(2).
01 WS-CHAR PIC X.
"#,
        r#"
    MOVE FUNCTION HEX-OF("A") TO WS-HEX.
    MOVE FUNCTION HEX-TO-CHAR("41") TO WS-CHAR.
"#,
    ));
}

#[test]
fn test_intrinsics_bit_operations() {
    compile_ok(&p(
        r#"
01 WS-BIT1 PIC X(8) VALUE "00001111".
01 WS-BIT2 PIC X(8) VALUE "01010101".
01 WS-RES PIC X(8).
01 WS-VAL PIC 9(3).
"#,
        r#"
    MOVE FUNCTION BIT-AND(WS-BIT1 WS-BIT2) TO WS-RES.
    MOVE FUNCTION BIT-OR(WS-BIT1 WS-BIT2) TO WS-RES.
    MOVE FUNCTION BIT-XOR(WS-BIT1 WS-BIT2) TO WS-RES.
    MOVE FUNCTION BIT-NOT(WS-BIT1) TO WS-RES.
    COMPUTE WS-VAL = FUNCTION INTEGER-OF-BOOLEAN(WS-BIT1).
"#,
    ));
}

#[test]
fn test_intrinsics_module_info() {
    compile_ok(&p(
        r#"
01 WS-NAME PIC X(30).
"#,
        r#"
    MOVE FUNCTION MODULE-NAME TO WS-NAME.
"#,
    ));
}
