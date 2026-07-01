use super::helpers::compile_ok_check;

fn build_move_program(pic: &str, value: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-FIELD PIC {pic} VALUE 0.\nPROCEDURE DIVISION.\n    MOVE {value} TO WS-FIELD.\n    STOP RUN."
    )
}

#[test]
fn test_move_numeric_literal_matrix() {
    let pics = [
        "9", "9(1)", "9(2)", "9(3)", "9(4)", "9(5)", "9(6)", "S9", "S9(1)", "S9(2)",
        "S9(3)", "S9(4)", "S9(5)", "S9(6)", "9(2)V9", "9(3)V99", "9(4)V9", "9(1)V9",
        "9(5)V99", "S9(3)V9",
    ];
    let values = [
        "0", "1", "-1", "2", "-2", "3", "-3", "4", "-4", "5", "-5", "6", "-6", "7", "-7",
        "8", "-8", "9", "-9", "10", "-10", "11", "-11", "12", "-12", "15", "-15", "20",
        "-20", "25", "-25", "30", "-30", "33", "-33", "50", "-50", "75", "-75", "100",
        "-100", "101", "-101", "123", "-123", "255", "-255", "999", "-999",
    ];

    let mut checked = 0;
    for pic in pics {
        for value in values {
            let src = build_move_program(pic, value);
            assert!(
                compile_ok_check(&src),
                "numeric move failed for pic {pic} with value {value}"
            );
            checked += 1;
        }
    }

    assert!(checked >= 1000, "expected at least 1000 move cases, got {checked}");
}
