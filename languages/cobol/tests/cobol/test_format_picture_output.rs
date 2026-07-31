use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn pic_z_suppresses_leading_zero() {
    let out = run_prints(&p(
        "01 N PIC Z(5) VALUE 42.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["   42"]);
}

#[test]
fn pic_z_all_zeros_blank() {
    let out = run_prints(&p(
        "01 N PIC Z(5) VALUE 0.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["     "]);
}

#[test]
fn pic_z_full_value_displayed() {
    let out = run_prints(&p(
        "01 N PIC Z(5) VALUE 99999.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["99999"]);
}

#[test]
fn pic_asterisk_fill_leading() {
    let out = run_prints(&p(
        "01 N PIC ***** VALUE 42.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["***42"]);
}

#[test]
fn pic_asterisk_all_zeros_all_stars() {
    let out = run_prints(&p(
        "01 N PIC ***** VALUE 0.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["*****"]);
}

#[test]
fn pic_plus_positive_shows_sign() {
    let out = run_prints(&p(
        "01 N PIC +9(4) VALUE 42.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["+0042"]);
}

#[test]
fn pic_plus_negative_shows_minus() {
    let out = run_prints(&p(
        "01 N PIC +9(4) VALUE -42.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["-0042"]);
}

#[test]
fn pic_minus_positive_shows_space() {
    let out = run_prints(&p(
        "01 N PIC -9(4) VALUE 42.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec![" 0042"]);
}

#[test]
fn pic_minus_negative_shows_minus() {
    let out = run_prints(&p(
        "01 N PIC -9(4) VALUE -42.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["-0042"]);
}

#[test]
fn pic_dot_inserts_decimal_point() {
    let out = run_prints(&p(
        "01 N PIC 9(4).99 VALUE 1234.56.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["1234.56"]);
}

#[test]
fn pic_comma_inserts_thousands() {
    let out = run_prints(&p(
        "01 N PIC 9,9(3) VALUE 1999.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["1,999"]);
}

#[test]
fn pic_b_inserts_blank() {
    let out = run_prints(&p(
        "01 N PIC 9(3)B9(3) VALUE 123456.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["123 456"]);
}

#[test]
fn pic_zero_inserts_literal_zero() {
    let out = run_prints(&p(
        "01 N PIC 9(3)09(2) VALUE 12345.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["1230045"]);
}

#[test]
fn pic_z_with_decimal() {
    let out = run_prints(&p(
        "01 N PIC ZZ9.99 VALUE 0.75.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["  0.75"]);
}

#[test]
fn pic_z_partial_suppression() {
    let out = run_prints(&p(
        "01 N PIC ZZ99 VALUE 0099.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["  99"]);
}

#[test]
fn pic_dollar_float_leading() {
    let out = run_prints(&p(
        "01 N PIC $$$9 VALUE 42.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["$$42"]);
}

#[test]
fn pic_cr_positive_blank() {
    let out = run_prints(&p(
        "01 N PIC 9(4)CR VALUE 100.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["0100  "]);
}

#[test]
fn pic_cr_negative_shows_cr() {
    let out = run_prints(&p(
        "01 N PIC 9(4)CR VALUE -100.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["0100CR"]);
}

#[test]
fn pic_db_positive_blank() {
    let out = run_prints(&p(
        "01 N PIC 9(4)DB VALUE 50.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["0050  "]);
}

#[test]
fn pic_db_negative_shows_db() {
    let out = run_prints(&p(
        "01 N PIC 9(4)DB VALUE -50.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["0050DB"]);
}

#[test]
fn pic_z_numeric_one() {
    let out = run_prints(&p(
        "01 N PIC Z(4) VALUE 1.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["   1"]);
}

#[test]
fn pic_comma_and_dot_combined() {
    let out = run_prints(&p(
        "01 N PIC 9,9(3).99 VALUE 1234.56.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["1,234.56"]);
}

#[test]
fn pic_plus_z_and_decimal() {
    compile_ok(&p(
        "01 N PIC +Z(4).99 VALUE -12.5.",
        "    DISPLAY N.",
    ));
}

#[test]
fn pic_star_with_decimal() {
    let out = run_prints(&p(
        "01 N PIC ***9.99 VALUE 1.25.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["***1.25"]);
}

#[test]
fn pic_edited_after_move() {
    let out = run_prints(&p(
        "01 SRC PIC 9(4) VALUE 1234.\n01 DST PIC ZZZ9 VALUE 0.",
        "    MOVE SRC TO DST.\n    DISPLAY DST.",
    ));
    assert_eq!(out, vec!["1234"]);
}

#[test]
fn pic_edited_zero_suppression_after_move() {
    let out = run_prints(&p(
        "01 SRC PIC 9(4) VALUE 5.\n01 DST PIC ZZZ9 VALUE 0.",
        "    MOVE SRC TO DST.\n    DISPLAY DST.",
    ));
    assert_eq!(out, vec!["   5"]);
}

#[test]
fn pic_b_separates_phone_parts() {
    let out = run_prints(&p(
        "01 PHONE PIC 9(3)BB9(3)BB9(4) VALUE 5551234567.",
        "    DISPLAY PHONE.",
    ));
    assert_eq!(out, vec!["555  123  4567"]);
}

#[test]
fn pic_zero_insert_in_date_format() {
    let out = run_prints(&p(
        "01 YYYYMMDD PIC 9(4)09(2)09(2) VALUE 20230615.",
        "    DISPLAY YYYYMMDD.",
    ));
    assert_eq!(out, vec!["2023006015"]);
}

#[test]
fn pic_plus_with_decimal_compiles() {
    compile_ok(&p(
        "01 N PIC +9(5).99 VALUE -123.45.",
        "    DISPLAY N.",
    ));
}
