use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn redefines_numeric_over_alpha() {
    let out = run_prints(&p(
        "01 BASE-FIELD PIC X(4) VALUE \"0042\".\n01 NUMERIC-VIEW REDEFINES BASE-FIELD PIC 9(4).",
        "    DISPLAY NUMERIC-VIEW.",
    ));
    assert_eq!(out, vec!["0042"]);
}

#[test]
fn redefines_alpha_over_numeric() {
    compile_ok(&p(
        "01 NUM-BASE PIC 9(4) VALUE 1234.\n01 ALPHA-VIEW REDEFINES NUM-BASE PIC X(4).",
        "    DISPLAY ALPHA-VIEW.",
    ));
}

#[test]
fn redefines_same_memory_different_pic() {
    let out = run_prints(&p(
        "01 BASE PIC X(4) VALUE \"1234\".\n01 RDEF REDEFINES BASE PIC 9(4).",
        "    MOVE 5678 TO RDEF.\n    DISPLAY BASE.",
    ));
    assert_eq!(out, vec!["5678"]);
}

#[test]
fn redefines_group_over_elementary() {
    let out = run_prints(&p(
        "01 COMBO PIC X(6) VALUE \"ABCDEF\".\n01 SPLIT REDEFINES COMBO.\n   05 PART1 PIC X(3).\n   05 PART2 PIC X(3).",
        "    DISPLAY PART1.\n    DISPLAY PART2.",
    ));
    assert_eq!(out, vec!["ABC", "DEF"]);
}

#[test]
fn redefines_elementary_over_group() {
    let out = run_prints(&p(
        "01 GRP.\n   05 GA PIC X(3) VALUE \"XYZ\".\n   05 GB PIC X(3) VALUE \"123\".\n01 COMBINED REDEFINES GRP PIC X(6).",
        "    DISPLAY COMBINED.",
    ));
    assert_eq!(out, vec!["XYZ123"]);
}

#[test]
fn redefines_numeric_group_parts() {
    let out = run_prints(&p(
        "01 DATE-FIELD PIC X(8) VALUE \"20231225\".\n01 DATE-PARTS REDEFINES DATE-FIELD.\n   05 DP-YEAR PIC 9(4).\n   05 DP-MONTH PIC 9(2).\n   05 DP-DAY PIC 9(2).",
        "    DISPLAY DP-YEAR.\n    DISPLAY DP-MONTH.\n    DISPLAY DP-DAY.",
    ));
    assert_eq!(out, vec!["2023", "12", "25"]);
}

#[test]
fn redefines_write_via_redefine_reads_original() {
    let out = run_prints(&p(
        "01 BUF PIC X(4) VALUE \"XXXX\".\n01 ALIAS REDEFINES BUF PIC X(4).",
        "    MOVE \"COBO\" TO ALIAS.\n    DISPLAY BUF.",
    ));
    assert_eq!(out, vec!["COBO"]);
}

#[test]
fn redefines_two_views_same_base() {
    let out = run_prints(&p(
        "01 BASE PIC X(2) VALUE \"AB\".\n01 VIEW-1 REDEFINES BASE PIC X(2).\n01 VIEW-2 REDEFINES BASE PIC X(2).",
        "    DISPLAY VIEW-1.\n    DISPLAY VIEW-2.",
    ));
    assert_eq!(out, vec!["AB", "AB"]);
}

#[test]
fn redefines_numeric_comp_over_alpha() {
    compile_ok(&p(
        "01 BUF PIC X(4) VALUE \"\\x00\\x00\\x00\\x01\".\n01 INT-VIEW REDEFINES BUF PIC 9(9) COMP.",
        "    DISPLAY INT-VIEW.",
    ));
}

#[test]
fn move_group_to_group_same_size() {
    let out = run_prints(&p(
        "01 SRC.\n   05 S1 PIC X(3) VALUE \"AAA\".\n   05 S2 PIC X(3) VALUE \"BBB\".\n01 DST.\n   05 D1 PIC X(3) VALUE \"XXX\".\n   05 D2 PIC X(3) VALUE \"YYY\".",
        "    MOVE SRC TO DST.\n    DISPLAY D1.\n    DISPLAY D2.",
    ));
    assert_eq!(out, vec!["AAA", "BBB"]);
}

#[test]
fn move_group_to_larger_group_truncates() {
    compile_ok(&p(
        "01 SRC.\n   05 S1 PIC X(4) VALUE \"ABCD\".\n01 DST.\n   05 D1 PIC X(4) VALUE \"XXXX\".\n   05 D2 PIC X(4) VALUE \"YYYY\".",
        "    MOVE SRC TO DST.",
    ));
}

#[test]
fn redefines_condition_name_on_redefine() {
    compile_ok(&p(
        "01 SWITCH PIC X VALUE \"Y\".\n01 SWITCH-NUM REDEFINES SWITCH PIC 9.\n    88 SWITCH-ON VALUE 1.",
        "    IF SWITCH-ON\n        DISPLAY \"ON\"\n    ELSE\n        DISPLAY \"OFF\"\n    END-IF.",
    ));
}

#[test]
fn redefines_three_level_group_redefine() {
    let out = run_prints(&p(
        "01 CODE-AREA PIC X(8) VALUE \"20230115\".\n01 PARSED REDEFINES CODE-AREA.\n   05 YEAR PIC 9(4).\n   05 MON  PIC 9(2).\n   05 DAY  PIC 9(2).",
        "    DISPLAY YEAR.\n    DISPLAY MON.\n    DISPLAY DAY.",
    ));
    assert_eq!(out, vec!["2023", "01", "15"]);
}

#[test]
fn redefines_char_level_access() {
    let out = run_prints(&p(
        "01 WORD PIC X(4) VALUE \"ABCD\".\n01 CHARS REDEFINES WORD.\n   05 C1 PIC X.\n   05 C2 PIC X.\n   05 C3 PIC X.\n   05 C4 PIC X.",
        "    DISPLAY C1.\n    DISPLAY C2.\n    DISPLAY C3.\n    DISPLAY C4.",
    ));
    assert_eq!(out, vec!["A", "B", "C", "D"]);
}

#[test]
fn redefines_modifying_subfield_updates_parent() {
    let out = run_prints(&p(
        "01 FULL-DATE PIC X(8) VALUE \"20000101\".\n01 DATE-REDEF REDEFINES FULL-DATE.\n   05 D-YEAR PIC 9(4).\n   05 D-REST PIC X(4).",
        "    MOVE 2024 TO D-YEAR.\n    DISPLAY FULL-DATE.",
    ));
    assert_eq!(out, vec!["20240101"]);
}

#[test]
fn redefines_signed_view_of_unsigned() {
    compile_ok(&p(
        "01 BALANCE PIC 9(6) VALUE 100000.\n01 SIGNED-BAL REDEFINES BALANCE PIC S9(6).",
        "    DISPLAY SIGNED-BAL.",
    ));
}

#[test]
fn move_group_sets_filler_areas() {
    let out = run_prints(&p(
        "01 PATTERN PIC X(6) VALUE \"XYZABC\".\n01 OVERLAY REDEFINES PATTERN.\n   05 FILLER PIC X(3).\n   05 SECOND  PIC X(3).",
        "    DISPLAY SECOND.",
    ));
    assert_eq!(out, vec!["ABC"]);
}

#[test]
fn redefines_numeric_then_arithmetic() {
    let out = run_prints(&p(
        "01 BUF PIC X(4) VALUE \"0099\".\n01 N REDEFINES BUF PIC 9(4).",
        "    ADD 1 TO N.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["0100"]);
}

#[test]
fn redefines_group_can_be_moved_to() {
    let out = run_prints(&p(
        "01 BASE PIC X(6) VALUE \"AABBCC\".\n01 REDEF REDEFINES BASE.\n   05 R1 PIC X(2).\n   05 R2 PIC X(2).\n   05 R3 PIC X(2).",
        "    MOVE \"XX\" TO R1.\n    DISPLAY BASE.",
    ));
    assert_eq!(out, vec!["XXBBCC"]);
}

#[test]
fn move_group_from_group_different_sizes_pads() {
    let out = run_prints(&p(
        "01 SRC.\n   05 A PIC X(3) VALUE \"ABC\".\n01 DST.\n   05 B PIC X(6) VALUE \"XXXXXX\".",
        "    MOVE SRC TO DST.\n    DISPLAY B.",
    ));
    assert_eq!(out, vec!["ABC   "]);
}

#[test]
fn redefines_level05_in_group_redefine() {
    compile_ok(&p(
        "01 RECORD-AREA.\n   05 AREA-DATA PIC X(10).\n   05 NUMERIC-DATA REDEFINES AREA-DATA PIC 9(10).",
        "    MOVE 1234567890 TO NUMERIC-DATA.\n    DISPLAY AREA-DATA.",
    ));
}

#[test]
fn redefines_pic_9_v99_overlay() {
    let out = run_prints(&p(
        "01 AMOUNT PIC X(7) VALUE \"0012345\".\n01 AMT-NUM REDEFINES AMOUNT PIC 9(5)V99.",
        "    DISPLAY AMT-NUM.",
    ));
    assert_eq!(out, vec!["001234500"]);
}

#[test]
fn redefines_three_alternatives() {
    compile_ok(&p(
        "01 UNION-AREA PIC X(8) VALUE SPACES.\n01 INT-UNION REDEFINES UNION-AREA PIC 9(8).\n01 DATE-UNION REDEFINES UNION-AREA.\n   05 D-YEAR PIC 9(4).\n   05 D-MMDD PIC 9(4).",
        "    MOVE 20230115 TO INT-UNION.\n    DISPLAY D-YEAR.",
    ));
}

#[test]
fn redefines_used_in_evaluate() {
    let out = run_prints(&p(
        "01 BASE PIC X(2) VALUE \"01\".\n01 CODE-NUM REDEFINES BASE PIC 9(2).",
        "    EVALUATE CODE-NUM\n        WHEN 1\n            DISPLAY \"ONE\"\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["ONE"]);
}

#[test]
fn redefines_in_condition() {
    let out = run_prints(&p(
        "01 STATUS-BYTE PIC X VALUE \"Y\".\n01 STATUS-NUM REDEFINES STATUS-BYTE PIC 9.",
        "    IF STATUS-BYTE = \"Y\"\n        DISPLAY \"YES\"\n    ELSE\n        DISPLAY \"NO\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["YES"]);
}

#[test]
fn redefines_preserves_original_value() {
    let out = run_prints(&p(
        "01 ORIG PIC X(5) VALUE \"HELLO\".\n01 ALSO-ORIG REDEFINES ORIG PIC X(5).",
        "    DISPLAY ORIG.\n    DISPLAY ALSO-ORIG.",
    ));
    assert_eq!(out, vec!["HELLO", "HELLO"]);
}

#[test]
fn move_group_to_alpha_then_display() {
    let out = run_prints(&p(
        "01 SRC.\n   05 PART-A PIC X(4) VALUE \"ABCD\".\n   05 PART-B PIC X(4) VALUE \"EFGH\".\n01 DST PIC X(8) VALUE SPACES.",
        "    MOVE SRC TO DST.\n    DISPLAY DST.",
    ));
    assert_eq!(out, vec!["ABCDEFGH"]);
}

#[test]
fn redefines_char_extract_via_group() {
    let out = run_prints(&p(
        "01 RECORD PIC X(10) VALUE \"HELLO-COBO\".\n01 WORDS REDEFINES RECORD.\n   05 WORD1 PIC X(5).\n   05 FILLER PIC X.\n   05 WORD2 PIC X(4).",
        "    DISPLAY WORD1.\n    DISPLAY WORD2.",
    ));
    assert_eq!(out, vec!["HELLO", "COBO"]);
}
