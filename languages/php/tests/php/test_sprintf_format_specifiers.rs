use super::helpers::run_prints;

// ── Binary, octal, hex specifiers ─────────────────────────────

#[test]
fn sprintf_binary_format() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%b', 255); "#),
        vec!["11111111"]
    );
}
#[test]
fn sprintf_octal_format() {
    assert_eq!(run_prints(r#"<?php echo sprintf('%o', 8); "#), vec!["10"]);
}
#[test]
fn sprintf_lowercase_hex() {
    assert_eq!(run_prints(r#"<?php echo sprintf('%x', 255); "#), vec!["ff"]);
}
#[test]
fn sprintf_uppercase_hex() {
    assert_eq!(run_prints(r#"<?php echo sprintf('%X', 255); "#), vec!["FF"]);
}
#[test]
fn sprintf_hex_with_prefix() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('0x%x', 255); "#),
        vec!["0xff"]
    );
}

// ── Width and padding ──────────────────────────────────────────

#[test]
fn sprintf_zero_padded_integer() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%05d', 42); "#),
        vec!["00042"]
    );
}
#[test]
fn sprintf_space_padded_integer() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%5d', 42); "#),
        vec!["   42"]
    );
}
#[test]
fn sprintf_left_aligned_string() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%-10s|', 'hi'); "#),
        vec!["hi        |"]
    );
}
#[test]
fn sprintf_right_aligned_string() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%10s|', 'hi'); "#),
        vec!["        hi|"]
    );
}
#[test]
fn sprintf_custom_pad_char() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%\'#10s', 'hi'); "#),
        vec!["########hi"]
    );
}

// ── Sign and precision ─────────────────────────────────────────

#[test]
fn sprintf_positive_sign() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%+d', 42); "#),
        vec!["+42"]
    );
}
#[test]
fn sprintf_negative_sign() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%+d', -42); "#),
        vec!["-42"]
    );
}
#[test]
fn sprintf_float_precision() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%.4f', M_PI); "#),
        vec!["3.1416"]
    );
}
#[test]
fn sprintf_scientific_notation() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%e', 123456.789); "#),
        vec!["1.234568e+5"]
    );
}
#[test]
fn sprintf_scientific_uppercase() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%E', 0.000123); "#),
        vec!["1.230000E-4"]
    );
}

// ── Argument swapping ──────────────────────────────────────────

#[test]
fn sprintf_argument_swap() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%2$s %1$s', 'world', 'hello'); "#),
        vec!["hello world"]
    );
}
#[test]
fn sprintf_argument_reuse() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%1$s %1$s', 'echo'); "#),
        vec!["echo echo"]
    );
}

// ── %u unsigned integer ────────────────────────────────────────

#[test]
fn sprintf_char_format() {
    assert_eq!(run_prints(r#"<?php echo sprintf('%c', 65); "#), vec!["A"]);
}
#[test]
fn sprintf_char_lowercase_z() {
    assert_eq!(run_prints(r#"<?php echo sprintf('%c', 122); "#), vec!["z"]);
}

// ── Multiple values ────────────────────────────────────────────

#[test]
fn sprintf_multiple_placeholders() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%s is %d years old', 'Alice', 30); "#),
        vec!["Alice is 30 years old"]
    );
}
#[test]
fn sprintf_mixed_types() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%d %s %.2f', 1, 'item', 9.99); "#),
        vec!["1 item 9.99"]
    );
}

// ── vsprintf ───────────────────────────────────────────────────

#[test]
fn vsprintf_with_array() {
    assert_eq!(
        run_prints(r#"<?php echo vsprintf('%s=%d', ['x', 42]); "#),
        vec!["x=42"]
    );
}

// ── sscanf ─────────────────────────────────────────────────────

#[test]
fn sscanf_extracts_values() {
    assert_eq!(
        run_prints(
            r#"<?php
[$a, $b] = sscanf("age 30", "age %d");
echo $a;
"#
        ),
        vec!["30"]
    );
}
#[test]
fn sscanf_string_and_int() {
    assert_eq!(
        run_prints(
            r#"<?php
[$name, $age] = sscanf("Alice 25", "%s %d");
echo "$name,$age";
"#
        ),
        vec!["Alice,25"]
    );
}

// ── number_format vs sprintf for floats ──────────────────────

#[test]
fn sprintf_g_removes_trailing_zeros() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%g', 100.0); "#),
        vec!["100"]
    );
}
#[test]
fn sprintf_f_preserves_trailing_zeros() {
    assert_eq!(
        run_prints(r#"<?php echo sprintf('%.3f', 1.1); "#),
        vec!["1.100"]
    );
}
