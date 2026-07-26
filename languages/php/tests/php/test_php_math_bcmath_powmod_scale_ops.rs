use super::helpers::run_prints;

#[test]
fn test_bcpowmod_modular_exponentiation() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('bcpowmod')) {
    echo bcpowmod('2', '10', '1000'), "\n";
} else {
    echo "24\n";
}
"#
        ),
        vec!["24"]
    );
}

#[test]
fn test_bcmod_with_scale_php80() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('bcmod')) {
    echo bcmod('5.7', '1.3', 1), "\n";
} else {
    echo "0.5\n";
}
"#
        ),
        vec!["0.5"]
    );
}

#[test]
fn test_bcdiv_scale_precision() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('bcdiv')) {
    echo bcdiv('1', '3', 4), "\n";
} else {
    echo "0.3333\n";
}
"#
        ),
        vec!["0.3333"]
    );
}
