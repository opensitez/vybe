//! `random_bytes`, `random_int`, and `openssl_random_pseudo_bytes` when available.

crate::php_cases! {
    random_int_within_inclusive_range => {
        r#"<?php
$n = random_int(10, 10);
echo $n;
"#,
        ["10"]
    };

    random_int_range_produces_value_between_bounds => {
        r#"<?php
$n = random_int(1, 6);
echo ($n >= 1 && $n <= 6) ? 'in' : 'out';
"#,
        ["in"]
    };

    random_bytes_length_matches_argument => {
        r#"<?php
echo strlen(random_bytes(16));
"#,
        ["16"]
    };

    random_bytes_different_on_two_calls_usually => {
        r#"<?php
$a = bin2hex(random_bytes(4));
$b = bin2hex(random_bytes(4));
echo strlen($a) === 8 && strlen($b) === 8 ? 'ok' : 'bad';
"#,
        ["ok"]
    };

    bin2hex_random_bytes_is_hex_string => {
        r#"<?php
$h = bin2hex(random_bytes(3));
echo ctype_xdigit($h) && strlen($h) === 6 ? 'hex' : 'no';
"#,
        ["hex"]
    };

    mt_rand_inclusive_bounds => {
        r#"<?php
mt_srand(42);
$n = mt_rand(5, 5);
echo $n;
"#,
        ["5"]
    };

    mt_getrandmax_positive => {
        r#"<?php
echo mt_getrandmax() > 0 ? 'pos' : 'zero';
"#,
        ["pos"]
    };

    uniqid_prefix_included_when_set => {
        r#"<?php
echo str_starts_with(uniqid('pid'), 'pid') ? 'pref' : 'no';
"#,
        ["pref"]
    };

    uniqid_more_entropy_longer_than_without => {
        r#"<?php
echo strlen(uniqid('', true)) > strlen(uniqid()) ? 'long' : 'short';
"#,
        ["long"]
    };
}
