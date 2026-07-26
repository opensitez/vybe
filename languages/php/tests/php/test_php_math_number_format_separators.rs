use super::helpers::run_prints;

#[test]
fn test_number_format_custom_decimal_and_thousands() {
    assert_eq!(
        run_prints(
            r#"<?php
echo number_format(1234567.89, 2, ',', ' '), "\n";
"#
        ),
        vec!["1 234 567,89"]
    );
}

#[test]
fn test_number_format_zero_decimals() {
    assert_eq!(
        run_prints(
            r#"<?php
echo number_format(1234567.89, 0, '.', ','), "\n";
"#
        ),
        vec!["1,234,568"]
    );
}
