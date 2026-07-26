use super::helpers::run_prints;

#[test]
fn test_fdiv_by_zero_positive() {
    assert_eq!(
        run_prints(
            r#"<?php
$res = fdiv(1.0, 0.0);
echo is_infinite($res) && $res > 0 ? 'INF' : 'other', "\n";
"#
        ),
        vec!["INF"]
    );
}

#[test]
fn test_fdiv_by_zero_negative() {
    assert_eq!(
        run_prints(
            r#"<?php
$res = fdiv(-5.0, 0.0);
echo is_infinite($res) && $res < 0 ? '-INF' : 'other', "\n";
"#
        ),
        vec!["-INF"]
    );
}

#[test]
fn test_fdiv_zero_by_zero_nan() {
    assert_eq!(
        run_prints(
            r#"<?php
$res = fdiv(0.0, 0.0);
echo is_nan($res) ? 'NAN' : 'other', "\n";
"#
        ),
        vec!["NAN"]
    );
}

#[test]
fn test_fdiv_normal_float_division() {
    assert_eq!(
        run_prints(
            r#"<?php
echo fdiv(10.0, 4.0), "\n";
"#
        ),
        vec!["2.5"]
    );
}
