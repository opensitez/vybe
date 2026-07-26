//! Bit shifts, rotation-style patterns, `decbin`/`bindec` — not `&|^~` in `test_bitwise_operators.rs`.

crate::php_cases! {
    left_shift_doubles_bits => {
        r#"<?php
echo 3 << 2;
"#,
        ["12"]
    };

    right_shift_halves_bits => {
        r#"<?php
echo 16 >> 2;
"#,
        ["4"]
    };

    left_shift_zero_yields_same => {
        r#"<?php
echo 5 << 0;
"#,
        ["5"]
    };

    right_shift_large_zero => {
        r#"<?php
echo 1 >> 10;
"#,
        ["0"]
    };

    xor_swap_without_temp => {
        r#"<?php
$a = 5;
$b = 9;
$a ^= $b;
$b ^= $a;
$a ^= $b;
echo $a . $b;
"#,
        ["95"]
    };

    bitwise_and_mask_low_nibble => {
        r#"<?php
echo 0xFF & 0x0F;
"#,
        ["15"]
    };

    bitwise_or_sets_flag => {
        r#"<?php
$flags = 1;
$flags |= 4;
echo $flags;
"#,
        ["5"]
    };

    bitwise_not_inverts => {
        r#"<?php
echo ~0 & 7;
"#,
        ["7"]
    };

    decbin_positive_integer => {
        r#"<?php
echo decbin(10);
"#,
        ["1010"]
    };

    bindec_binary_string => {
        r#"<?php
echo bindec('101');
"#,
        ["5"]
    };

    dechex_uppercase_a => {
        r#"<?php
echo dechex(10);
"#,
        ["a"]
    };

    hexdec_lowercase => {
        r#"<?php
echo hexdec('ff');
"#,
        ["255"]
    };

    decoct_eight => {
        r#"<?php
echo decoct(8);
"#,
        ["10"]
    };

    octdec_literal => {
        r#"<?php
echo octdec('17');
"#,
        ["15"]
    };

    intdiv_floor_division => {
        r#"<?php
echo intdiv(7, 2);
"#,
        ["3"]
    };

    intdiv_negative => {
        r#"<?php
echo intdiv(-7, 2);
"#,
        ["-3"]
    };

    intdiv_modulo_consistency => {
        r#"<?php
$a = 17;
$b = 5;
echo intdiv($a, $b) * $b + ($a % $b);
"#,
        ["17"]
    };

    pow_integer_exponent => {
        r#"<?php
echo 2 ** 8;
"#,
        ["256"]
    };

    pow_zero_exponent => {
        r#"<?php
echo 9 ** 0;
"#,
        ["1"]
    };

    shift_assign_left => {
        r#"<?php
$n = 1;
$n <<= 3;
echo $n;
"#,
        ["8"]
    };

    shift_assign_right => {
        r#"<?php
$n = 32;
$n >>= 2;
echo $n;
"#,
        ["8"]
    };

    bitwise_and_assign => {
        r#"<?php
$n = 0b1111;
$n &= 0b1010;
echo $n;
"#,
        ["10"]
    };

    bitwise_or_assign => {
        r#"<?php
$n = 0b1000;
$n |= 0b0011;
echo $n;
"#,
        ["11"]
    };

    bitwise_xor_assign => {
        r#"<?php
$n = 0b1100;
$n ^= 0b1010;
echo $n;
"#,
        ["6"]
    };

    pack_unpack_unsigned_short => {
        r#"<?php
$bytes = pack('n', 258);
echo unpack('n', $bytes)[1];
"#,
        ["258"]
    };

    shift_after_addition => {
        r#"<?php
echo 1 + 2 << 2;
echo '|';
echo (1 + 2) << 2;
"#,
        ["8|12"]
    };

    shift_before_addition => {
        r#"<?php
echo 1 << 2 + 1;
echo '|';
echo (1 << 2) + 1;
"#,
        ["8|5"]
    };

    xor_swap_no_temp_with_precedence => {
        r#"<?php
$a = 3;
$b = 10;
$a ^= $b;
$b ^= $a;
$a ^= $b;
echo $a;
echo $b;
"#,
        ["1010"]
    };
}
