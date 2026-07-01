use super::helpers::compile_ok_check;

fn make_program(pic_a: &str, pic_b: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-BASE PIC {pic_a} VALUE \"ABCD\".\n01 WS-ALIAS REDEFINES WS-BASE PIC {pic_b}.\nPROCEDURE DIVISION.\n    DISPLAY WS-ALIAS.\n    STOP RUN."
    )
}

#[test]
fn test_redefines_matrix() {
    let cases = [
        ("X(4)", "X(4)"),
        ("X(4)", "X(2)"),
        ("X(4)", "9(4)"),
        ("X(6)", "X(6)"),
        ("X(6)", "X(3)"),
        ("X(6)", "9(6)"),
        ("9(4)", "9(4)"),
        ("9(4)", "X(4)"),
        ("9(4)", "X(2)"),
        ("9(6)", "9(3)"),
    ];

    for (base, alias) in cases {
        assert!(compile_ok_check(&make_program(base, alias)), "redefines failed for {base} -> {alias}");
    }
}
