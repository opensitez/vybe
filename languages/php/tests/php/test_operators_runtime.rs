//! Compound assignment, null coalescing assignment, exponent, and silence operator.

crate::php_cases! {
    concat_assign_appends_string => {
        r#"<?php
$s = 'a';
$s .= 'b';
echo $s;
"#,
        ["ab"]
    };

    plus_assign_accumulates_integer => {
        r#"<?php
$n = 1;
$n += 4;
echo $n;
"#,
        ["5"]
    };

    null_coalesce_assign_sets_when_null => {
        r#"<?php
$a = null;
$a ??= 'default';
echo $a;
"#,
        ["default"]
    };

    null_coalesce_assign_skips_when_set => {
        r#"<?php
$a = 'keep';
$a ??= 'replace';
echo $a;
"#,
        ["keep"]
    };

    exponent_assign_squares_in_place => {
        r#"<?php
$n = 3;
$n **= 2;
echo $n;
"#,
        ["9"]
    };

    bitwise_or_assign_sets_flag => {
        r#"<?php
$m = 0b001;
$m |= 0b010;
echo decbin($m);
"#,
        ["11"]
    };

    bitwise_and_assign_clears_bits => {
        r#"<?php
$m = 0b111;
$m &= 0b101;
echo decbin($m);
"#,
        ["101"]
    };

    xor_assign_toggles_bits => {
        r#"<?php
$m = 0b1010;
$m ^= 0b1100;
echo decbin($m);
"#,
        ["110"]
    };

    shift_left_assign_doubles => {
        r#"<?php
$n = 3;
$n <<= 1;
echo $n;
"#,
        ["6"]
    };

    shift_right_assign_halves => {
        r#"<?php
$n = 8;
$n >>= 1;
echo $n;
"#,
        ["4"]
    };

    silence_operator_suppresses_undefined_variable_notice => {
        r#"<?php
echo @$missing ?? 'fallback';
"#,
        ["fallback"]
    };

    ternary_short_form_selects_truthy_branch => {
        r#"<?php
echo 1 ? 'yes' : 'no';
"#,
        ["yes"]
    };

    elvis_operator_returns_first_non_empty => {
        r#"<?php
echo '' ?: 'fallback';
"#,
        ["fallback"]
    };

    identity_vs_equality_int_and_string => {
        r#"<?php
echo (1 == '1') ? 'eq' : 'ne';
echo (1 === '1') ? 'id' : 'nid';
"#,
        ["eqnid"]
    };

    modulo_assign_wraps_counter => {
        r#"<?php
$n = 7;
$n %= 4;
echo $n;
"#,
        ["3"]
    };
}
