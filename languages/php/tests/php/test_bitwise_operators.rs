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

#[test]
fn shift_parentheses_precedence_runtime() {
    assert_eq!(run_prints(r#"<?php echo (1 << 2) + 1; "#), vec!["5"]);
}

#[test]
fn shift_with_zero_amount_runtime() {
    assert_eq!(
        run_prints(r#"<?php echo (3 << 0); echo '|'; echo (3 >> 0); "#),
        vec!["3|3"]
    );
}

#[test]
fn bitmask_with_multiple_flips_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$flags = 0b1010;
$flags ^= 0b0011;
$flags |= 0b0100;
$flags &= 0b1111;
echo decbin($flags);
"#
        ),
        vec!["1101"]
    );
}

#[test]
fn bitwise_assignment_chain_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$v = 12;
$v &= 14;
$v |= 1;
$v ^= 3;
echo $v;
"#
        ),
        vec!["14"]
    );
}

#[test]
fn unpack_unsigned_16bit_big_and_small_endian_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$b = pack('n', 513);
echo ord($b[0]);
echo '|';
$u = unpack('n', $b);
echo $u[1];
"#
        ),
        vec!["2|513"]
    );
}

#[test]
fn unary_bitwise_parenthesis_mix_runtime() {
    assert_eq!(run_prints(r#"<?php echo ~(-1) & 0xFF; "#), vec!["0"]);
}

#[test]
fn bitwise_string_operands_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo ord(('a' | 'b')[0]) . '|';
echo ord(('c' & 'b')[0]) . '|';
echo ord(('a' ^ 'a')[0]);
"#
        ),
        vec!["99|98|0"]
    );
}

#[test]
fn bitwise_shift_then_bitwise_or_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo (1 << 3) | 2;
echo '|';
echo (1 | 1 << 3);
"#
        ),
        vec!["10|9"]
    );
}

#[test]
fn bitwise_precedence_chain_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo 7 & 3 | 1;
echo '|';
echo 7 ^ 3 & 1;
echo '|';
echo 7 | 3 ^ 1;
            "#
        ),
        vec!["3|6|7"]
    );
}

#[test]
fn bitwise_parentheses_override_precedence_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo (7 & 3) | 1;
echo '|';
echo 7 & (3 | 1);
echo '|';
echo (7 ^ 3) & 1;
            "#
        ),
        vec!["3|3|0"]
    );
}

#[test]
fn shift_with_boolean_and_arithmetic_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo (true ? 1 : 0) << 2;
echo '|';
echo (false ? 1 : 2) >> 1;
echo '|';
echo (1 + 2) << 2 + 1;
"#
        ),
        vec!["4|1|24"]
    );
}

#[test]
fn bitwise_not_numeric_and_string_mix_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo ~(-1) . '|';
echo (~0) . '|';
echo ord(~'A');
"#
        ),
        vec!["0|-1|190"]
    );
}

#[test]
fn bitwise_negative_shift_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo (-8 >> 1) . '|';
echo (-8 << 1) . '|';
echo (-1 >> 0);
            "#
        ),
        vec!["-4|-16|-1"]
    );
}

#[test]
fn bitwise_shift_and_boolean_truthiness_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo (1 << 2) && 1 ? '1' : '0';
echo '|';
echo (0 << 2) ? '1' : '0';
echo '|';
echo (1 >> 2) ? '1' : '0';
            "#
        ),
        vec!["1|0|0"]
    );
}

#[test]
fn bitwise_chained_flags_state_mutation_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$state = 0;
$state |= 0b0001;
$state |= 0b0100;
echo decbin($state);
$state &= ~0b0001;
echo '|';
echo decbin($state);
$state ^= 0b0110;
echo '|';
echo decbin($state);
"#
        ),
        vec!["101|100|10"]
    );
}
