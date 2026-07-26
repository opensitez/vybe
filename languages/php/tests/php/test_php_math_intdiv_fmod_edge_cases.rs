use super::helpers::run_prints;

#[test]
fn test_intdiv_positive_numbers() {
    assert_eq!(
        run_prints(
            r#"<?php
echo intdiv(10, 3), "\n";
"#
        ),
        vec!["3"]
    );
}

#[test]
fn test_intdiv_negative_numerator() {
    assert_eq!(
        run_prints(
            r#"<?php
echo intdiv(-10, 3) . ',' . intdiv(10, -3) . ',' . intdiv(-10, -3), "\n";
"#
        ),
        vec!["-3,-3,3"]
    );
}

#[test]
fn test_intdiv_division_by_zero_throws() {
    assert_eq!(
        run_prints(
            r#"<?php
try {
    intdiv(5, 0);
    echo "no_error\n";
} catch (DivisionByZeroError $e) {
    echo "div_zero_error\n";
}
"#
        ),
        vec!["div_zero_error"]
    );
}

#[test]
fn test_fmod_floating_remainder() {
    assert_eq!(
        run_prints(
            r#"<?php
echo fmod(5.7, 1.3), "\n";
"#
        ),
        vec!["0.5"]
    );
}

#[test]
fn test_fmod_negative_arguments() {
    assert_eq!(
        run_prints(
            r#"<?php
echo fmod(-5.7, 1.3) . ',' . fmod(5.7, -1.3), "\n";
"#
        ),
        vec!["-0.5,0.5"]
    );
}
