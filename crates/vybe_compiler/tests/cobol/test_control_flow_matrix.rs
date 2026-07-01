use super::helpers::compile_ok_check;

fn make_program(cond: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC 9 VALUE 1.\n01 WS-B PIC 9 VALUE 2.\n01 WS-C PIC 9 VALUE 3.\nPROCEDURE DIVISION.\n    IF {cond}\n        DISPLAY \"YES\"\n    ELSE\n        DISPLAY \"NO\"\n    END-IF.\n    STOP RUN."
    )
}

#[test]
fn test_condition_matrix() {
    let conditions = [
        "WS-A > 0", "WS-A = 1", "WS-A NOT = 0", "WS-A < 2", "WS-A <= 1", "WS-A >= 1", "WS-A > 0 AND WS-B > 1",
        "WS-A > 0 AND WS-B > 3", "WS-A > 0 OR WS-B > 3", "WS-A > 0 OR WS-B < 0", "NOT WS-A = 0",
        "NOT (WS-A = 0)", "WS-A > 0 AND (WS-B > 1 OR WS-C > 10)", "WS-A > 0 AND (WS-B > 1 AND WS-C > 2)",
        "WS-A > 0 OR (WS-B > 10 AND WS-C > 10)", "WS-A = 1 AND WS-B = 2 AND WS-C = 3", "WS-A = 1 AND WS-B = 2 AND WS-C = 4",
        "WS-A < 2 AND WS-B < 3", "WS-A <= 1 AND WS-B <= 2", "WS-A >= 1 AND WS-B >= 2", "WS-A > 0 AND WS-B > 1 AND WS-C > 2",
        "WS-A = 1 OR WS-B = 5", "WS-A = 5 OR WS-B = 2", "WS-A NOT > 0", "WS-A NOT < 0", "WS-A NOT <= 0",
        "WS-A NOT >= 2", "WS-A NOT = 2", "WS-A > 0 AND WS-B > 0 AND WS-C > 0", "WS-A > 0 OR WS-B > 0 OR WS-C > 0",
        "WS-A = 1 OR WS-B = 2 OR WS-C = 3", "WS-A = 1 AND WS-B = 3", "WS-A = 1 OR WS-B = 3", "WS-A = 0 OR WS-B = 2",
    ];

    for cond in conditions {
        assert!(compile_ok_check(&make_program(cond)), "condition failed: {cond}");
    }
}
