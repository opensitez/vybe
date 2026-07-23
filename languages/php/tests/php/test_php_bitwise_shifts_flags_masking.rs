use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Bitwise Operations & Flags Masking — &, |, ^, ~, <<, >>, status bitmasks, permission masks
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_bitwise_and_or_xor_not_primitives() {
    let out = run_prints(
        r#"<?php
$a = 0b1100; // 12
$b = 0b1010; // 10

$and = $a & $b; // 1000 (8)
$or = $a | $b;  // 1110 (14)
$xor = $a ^ $b; // 0110 (6)

echo "$and | $or | $xor";
"#,
    );
    assert_eq!(out, vec!["8 | 14 | 6"]);
}

#[test]
fn test_php_bitwise_left_right_shifts() {
    let out = run_prints(
        r#"<?php
$val = 4;
$shl = $val << 2; // 16
$shr = $shl >> 3; // 2

echo "$shl | $shr";
"#,
    );
    assert_eq!(out, vec!["16 | 2"]);
}

#[test]
fn test_php_bitwise_permission_masking_pattern() {
    let out = run_prints(
        r#"<?php
const PERM_READ = 1 << 0;  // 1
const PERM_WRITE = 1 << 1; // 2
const PERM_EXEC = 1 << 2;  // 4

$userPerms = PERM_READ | PERM_EXEC;

echo ($userPerms & PERM_READ ? "1" : "0");
echo ($userPerms & PERM_WRITE ? "1" : "0");
echo ($userPerms & PERM_EXEC ? "1" : "0");
"#,
    );
    assert_eq!(out, vec!["101"]);
}

#[test]
fn test_php_bitwise_string_and_or_xor() {
    let out = run_prints(
        r#"<?php
$s1 = "A"; // ASCII 65 (01000001)
$s2 = " "; // ASCII 32 (00100000)
$res = $s1 | $s2; // ASCII 97 ('a')
echo $res;
"#,
    );
    assert_eq!(out, vec!["a"]);
}

#[test]
fn test_php_bitwise_not_complement() {
    compile_ok(
        r#"<?php
$x = 0;
$inv = ~$x;
echo is_int($inv) ? "INT_NOT_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_bitwise_compound_assignment_operators() {
    compile_ok(
        r#"<?php
$flags = 0;
$flags |= (1 << 0);
$flags |= (1 << 1);
$flags &= ~(1 << 0);
echo $flags === 2 ? "COMPOUND_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_bitwise_shift_overflow_behavior() {
    compile_ok(
        r#"<?php
$val = 1 << 30;
echo is_int($val) ? "INT_SHIFT" : "FLOAT_SHIFT";
"#,
    );
}

#[test]
fn test_php_bitwise_mask_combining_options() {
    compile_ok(
        r#"<?php
$options = JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES | JSON_THROW_ON_ERROR;
echo is_int($options) ? "MASK_INT" : "FAIL";
"#,
    );
}

#[test]
fn test_php_bitwise_toggle_flag_xor() {
    compile_ok(
        r#"<?php
$flag = 0b0010;
$flag ^= 0b0010; // toggle off -> 0
echo $flag === 0 ? "TOGGLE_OFF" : "FAIL";
"#,
    );
}

#[test]
fn test_php_bitwise_string_xor_encryption() {
    compile_ok(
        r#"<?php
$text = "Hello";
$key = "K";
$encrypted = $text ^ str_repeat($key, strlen($text));
$decrypted = $encrypted ^ str_repeat($key, strlen($text));
echo $decrypted === $text ? "XOR_ROUNDTRIP_OK" : "FAIL";
"#,
    );
}
