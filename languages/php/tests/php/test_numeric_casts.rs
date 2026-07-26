//! Numeric string parsing edges — partial parses, invalid bases, and validators.

crate::php_cases! {
    hexdec_stops_at_first_non_hex_digit => {
        r#"<?php
echo hexdec('12zz');
"#,
        ["18"]
    };

    bindec_ignores_trailing_non_binary_digits => {
        r#"<?php
echo bindec('1012');
"#,
        ["5"]
    };

    octdec_stops_at_invalid_octal_digit => {
        r#"<?php
echo octdec('778');
"#,
        ["63"]
    };

    base_convert_from_hex_to_decimal => {
        r#"<?php
echo base_convert('ff', 16, 10);
"#,
        ["255"]
    };

    base_convert_invalid_from_base_throws_value_error => {
        r#"<?php
try { base_convert('10', 1, 10); echo 'ok'; }
catch (ValueError $e) { echo 'base'; }
"#,
        ["base"]
    };

    base_convert_invalid_to_base_throws_value_error => {
        r#"<?php
try { base_convert('10', 10, 37); echo 'ok'; }
catch (ValueError $e) { echo 'to-base'; }
"#,
        ["to-base"]
    };

    intval_stops_at_first_non_digit_in_string => {
        r#"<?php
echo intval('42px');
"#,
        ["42"]
    };

    intval_binary_prefixed_string => {
        r#"<?php
echo intval('0b1010');
"#,
        ["0"]
    };

    floatval_parses_leading_numeric_portion => {
        r#"<?php
echo floatval('3.14abc');
"#,
        ["3.14"]
    };

    is_numeric_rejects_two_decimal_points => {
        r#"<?php
echo is_numeric('12.3.4') ? 'yes' : 'no';
"#,
        ["no"]
    };

    is_numeric_accepts_scientific_notation => {
        r#"<?php
echo is_numeric('1e3') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    ctype_digit_rejects_negative_sign => {
        r#"<?php
echo ctype_digit('-3') ? 'yes' : 'no';
"#,
        ["no"]
    };

    ctype_digit_accepts_ascii_digits_only => {
        r#"<?php
echo ctype_digit('903') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    bcdiv_by_zero_returns_null_without_scale => {
        r#"<?php
$v = bcdiv('1', '0');
echo $v === null ? 'null' : (string)$v;
"#,
        ["null"]
    };

    bcmod_by_zero_returns_null => {
        r#"<?php
$v = bcmod('5', '0');
echo $v === null ? 'null' : (string)$v;
"#,
        ["null"]
    };

    bcscale_negative_throws_value_error => {
        r#"<?php
try { bcscale(-1); echo 'ok'; }
catch (ValueError $e) { echo 'scale'; }
"#,
        ["scale"]
    };

    number_format_negative_decimals_throws_value_error => {
        r#"<?php
try { number_format(1.2, -1); echo 'ok'; }
catch (ValueError $e) { echo 'decimals'; }
"#,
        ["decimals"]
    };

    round_halves_to_even_bankers_rounding => {
        r#"<?php
echo (string)round(2.5, 0, PHP_ROUND_HALF_EVEN);
"#,
        ["2"]
    };

    fmod_preserves_fractional_remainder_sign => {
        r#"<?php
echo fmod(5.5, 2);
"#,
        ["1.5"]
    };

    deg2rad_converts_180_to_pi_approx => {
        r#"<?php
echo deg2rad(180) > 3.14 ? 'pi' : 'small';
"#,
        ["pi"]
    };

    intval_parses_whitespace_and_sign => {
        r#"<?php
echo intval('  42');
echo '|';
echo intval(" \t -17");
echo '|';
echo intval(" +12");
"#,
        ["42|-17|12"]
    };

    intval_with_explicit_base => {
        r#"<?php
echo intval('10', 2);
echo '|';
echo intval('10', 8);
echo '|';
echo intval('10', 16);
echo '|';
echo intval('0b101', 2);
"#,
        ["2|8|16|5"]
    };

    intval_autodetect_base => {
        r#"<?php
echo intval('0x10', 0);
echo '|';
echo intval('010', 0);
"#,
        ["16|8"]
    };

    numeric_string_strictness => {
        r#"<?php
echo is_numeric('1_000') ? 'yes' : 'no';
echo '|';
echo is_numeric('0x10') ? 'yes' : 'no';
echo '|';
echo is_numeric('0') ? 'yes' : 'no';
echo '|';
echo is_numeric('  ') ? 'yes' : 'no';
"#,
        ["no|no|yes|no"]
    };
}
