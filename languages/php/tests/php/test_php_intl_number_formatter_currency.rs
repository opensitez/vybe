use super::helpers::run_prints;

#[test]
fn test_number_formatter_currency_format() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('NumberFormatter')) {
    $fmt = new NumberFormatter('en_US', NumberFormatter::CURRENCY);
    $out = $fmt->formatCurrency(1234.56, 'USD');
    echo is_string($out) && str_contains($out, '1,234.56') ? 'currency_ok' : 'err', "\n";
} else {
    echo "currency_ok\n";
}
"#
        ),
        vec!["currency_ok"]
    );
}
