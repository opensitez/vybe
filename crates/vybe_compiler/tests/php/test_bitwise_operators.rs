use super::helpers::run_prints;

// ── Bitwise AND / OR / XOR / NOT ─────────────────────────────

#[test]
fn bitwise_and() {
    assert_eq!(run_prints(r#"<?php echo 0b1100 & 0b1010; "#), vec!["8"]);
}
#[test]
fn bitwise_or() {
    assert_eq!(run_prints(r#"<?php echo 0b1100 | 0b1010; "#), vec!["14"]);
}
#[test]
fn bitwise_xor() {
    assert_eq!(run_prints(r#"<?php echo 0b1100 ^ 0b1010; "#), vec!["6"]);
}
#[test]
fn bitwise_not() {
    assert_eq!(run_prints(r#"<?php echo ~5; "#), vec!["-6"]);
}
#[test]
fn bitwise_not_zero() {
    assert_eq!(run_prints(r#"<?php echo ~0; "#), vec!["-1"]);
}

// ── Bit shifts ────────────────────────────────────────────────

#[test]
fn left_shift() {
    assert_eq!(run_prints(r#"<?php echo 1 << 4; "#), vec!["16"]);
}
#[test]
fn right_shift() {
    assert_eq!(run_prints(r#"<?php echo 64 >> 3; "#), vec!["8"]);
}
#[test]
fn left_shift_multiply_by_two() {
    assert_eq!(run_prints(r#"<?php echo 7 << 1; "#), vec!["14"]);
}
#[test]
fn right_shift_divide_by_two() {
    assert_eq!(run_prints(r#"<?php echo 20 >> 1; "#), vec!["10"]);
}

// ── Bitmask flags pattern ─────────────────────────────────────

#[test]
fn bitmask_flag_set() {
    assert_eq!(
        run_prints(
            r#"<?php
define('READ',  0b001);
define('WRITE', 0b010);
define('EXEC',  0b100);
$perms = READ | WRITE;
echo ($perms & READ)  ? 'r' : '-';
echo ($perms & WRITE) ? 'w' : '-';
echo ($perms & EXEC)  ? 'x' : '-';
"#
        ),
        vec!["rw-"]
    );
}
#[test]
fn bitmask_flag_toggle() {
    assert_eq!(
        run_prints(
            r#"<?php
$flags = 0b0101;
$flags ^= 0b0100;
echo decbin($flags);
"#
        ),
        vec!["1"]
    );
}
#[test]
fn bitmask_flag_clear() {
    assert_eq!(
        run_prints(
            r#"<?php
$flags = 0b1111;
$flags &= ~0b0010;
echo decbin($flags);
"#
        ),
        vec!["1101"]
    );
}

// ── Bitwise compound assignment ───────────────────────────────

#[test]
fn bitwise_and_assign() {
    assert_eq!(
        run_prints(r#"<?php $v = 0xFF; $v &= 0x0F; echo $v; "#),
        vec!["15"]
    );
}
#[test]
fn bitwise_or_assign() {
    assert_eq!(
        run_prints(r#"<?php $v = 0; $v |= 1; $v |= 4; echo $v; "#),
        vec!["5"]
    );
}
#[test]
fn bitwise_xor_assign() {
    assert_eq!(
        run_prints(r#"<?php $v = 0b1010; $v ^= 0b1111; echo decbin($v); "#),
        vec!["101"]
    );
}
#[test]
fn shift_assign_left() {
    assert_eq!(
        run_prints(r#"<?php $v = 1; $v <<= 8; echo $v; "#),
        vec!["256"]
    );
}
#[test]
fn shift_assign_right() {
    assert_eq!(
        run_prints(r#"<?php $v = 256; $v >>= 4; echo $v; "#),
        vec!["16"]
    );
}

// ── pack / unpack (binary data) ───────────────────────────────

#[test]
fn pack_unsigned_char() {
    assert_eq!(
        run_prints(r#"<?php $b = pack('C', 65); echo $b; "#),
        vec!["A"]
    );
}
#[test]
fn unpack_unsigned_char() {
    assert_eq!(
        run_prints(r#"<?php $arr = unpack('C*', 'ABC'); echo implode(',', $arr); "#),
        vec!["65,66,67"]
    );
}
#[test]
fn pack_unsigned_32bit() {
    assert_eq!(
        run_prints(r#"<?php $b = pack('N', 1); echo strlen($b); "#),
        vec!["4"]
    );
}

// ── Integer literals in different bases ───────────────────────

#[test]
fn hex_literal() {
    assert_eq!(run_prints(r#"<?php echo 0xFF; "#), vec!["255"]);
}
#[test]
fn octal_literal() {
    assert_eq!(run_prints(r#"<?php echo 0777; "#), vec!["511"]);
}
#[test]
fn binary_literal() {
    assert_eq!(run_prints(r#"<?php echo 0b11111111; "#), vec!["255"]);
}
#[test]
fn underscore_numeric_separator() {
    assert_eq!(run_prints(r#"<?php echo 1_000_000; "#), vec!["1000000"]);
}
