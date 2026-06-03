use super::helpers::{compile_ok, run_prints};

// ── sprintf basic format specifiers ──────────────────────────

#[test]
fn sprintf_string() {
    compile_ok(
        r#"<?php
echo sprintf("Hello, %s!", "World");
echo sprintf("%s and %s", "foo", "bar");
"#,
    );
}

#[test]
fn sprintf_integer() {
    compile_ok(
        r#"<?php
echo sprintf("%d", 42);
echo sprintf("%d + %d = %d", 3, 4, 7);
echo sprintf("%d", -99);
"#,
    );
}

#[test]
fn sprintf_float() {
    compile_ok(
        r#"<?php
echo sprintf("%f", 3.14);
echo sprintf("%.2f", 3.14159);
echo sprintf("%.4f", 1.0/3.0);
"#,
    );
}

#[test]
fn sprintf_padding_integer() {
    compile_ok(
        r#"<?php
echo sprintf("%05d", 42);      // 00042
echo sprintf("%-5d|", 42);     // 42   |
echo sprintf("%+d", 42);       // +42
echo sprintf("%+d", -42);      // -42
"#,
    );
}

#[test]
fn sprintf_padding_string() {
    compile_ok(
        r#########"<?php
echo sprintf("%10s", "hello");  // "     hello"
echo sprintf("%-10s|", "hi");   // "hi        |"
echo sprintf("%'#10s", "ok");   // "########ok"
"#########,
    );
}

#[test]
fn sprintf_hex_octal_binary() {
    compile_ok(
        r#"<?php
echo sprintf("%x", 255);   // ff
echo sprintf("%X", 255);   // FF
echo sprintf("%o", 8);     // 10
echo sprintf("%b", 10);    // 1010
echo sprintf("%08b", 10);  // 00001010
"#,
    );
}

#[test]
fn sprintf_scientific() {
    compile_ok(
        r#"<?php
echo sprintf("%e", 123456.789);  // 1.234568e+5
echo sprintf("%E", 0.000123);    // 1.230000E-4
echo sprintf("%.2e", 1234.5);
"#,
    );
}

#[test]
fn sprintf_unsigned() {
    compile_ok(
        r#"<?php
echo sprintf("%u", 42);
echo sprintf("%u", PHP_INT_MAX);
"#,
    );
}

#[test]
fn sprintf_char() {
    compile_ok(
        r#"<?php
echo sprintf("%c", 65);   // A
echo sprintf("%c", 97);   // a
echo sprintf("%c%c%c", 72, 105, 33);  // Hi!
"#,
    );
}

// ── Argument swapping ─────────────────────────────────────────

#[test]
fn sprintf_argument_swap() {
    assert_eq!(
        run_prints(
            r#"<?php
echo sprintf('%2$s %1$s', 'World', 'Hello');
echo sprintf('%1$s has %2$d items at $%3$.2f each', 'Cart', 3, 9.99);
"#
        ),
        vec![
            "Hello World".to_string(),
            "Cart has 3 items at $9.99 each".to_string(),
        ]
    );
}

#[test]
fn sprintf_argument_swap_repeat() {
    assert_eq!(
        run_prints(
            r#"<?php
echo sprintf('%1$s %1$s %2$s', 'la', 'land');
"#
        ),
        vec!["la la land".to_string()]
    );
}

// ── Width and precision ───────────────────────────────────────

#[test]
fn sprintf_width_precision_float() {
    compile_ok(
        r#"<?php
echo sprintf('%10.2f', 3.14);   // "      3.14"
echo sprintf('%-10.2f|', 3.14); // "3.14      |"
echo sprintf('%010.2f', 3.14);  // "0000003.14"
"#,
    );
}

#[test]
fn sprintf_width_dynamic() {
    compile_ok(
        r#"<?php
echo sprintf('%*d', 5, 42);   // PHP uses %5d style; *-width is non-standard but %5d works
echo sprintf('%5d', 42);
"#,
    );
}

// ── printf and vprintf ────────────────────────────────────────

#[test]
fn printf_basic() {
    compile_ok(
        r#"<?php
$written = printf("Name: %s, Age: %d\n", "Alice", 30);
echo $written > 0 ? 'wrote bytes' : 'nothing written';
"#,
    );
}

#[test]
fn vprintf_basic() {
    compile_ok(
        r#"<?php
$args = ["Bob", 25];
$written = vprintf("Name: %s, Age: %d\n", $args);
echo $written > 0 ? 'wrote bytes' : 'nothing';
"#,
    );
}

#[test]
fn vsprintf_basic() {
    compile_ok(
        r#"<?php
$args = ['PHP', '8.3'];
$result = vsprintf("%s version %s", $args);
echo $result;
"#,
    );
}

// ── sscanf ────────────────────────────────────────────────────

#[test]
fn sscanf_basic() {
    compile_ok(
        r#"<?php
$result = sscanf("Age: 25", "Age: %d");
echo $result[0];
"#,
    );
}

#[test]
fn sscanf_multiple() {
    compile_ok(
        r#"<?php
[$y, $m, $d] = sscanf("2024-01-15", "%d-%d-%d");
echo "$y-$m-$d";
"#,
    );
}

#[test]
fn sscanf_string_and_int() {
    compile_ok(
        r#"<?php
[$name, $age] = sscanf("Alice 30", "%s %d");
echo "$name is $age";
"#,
    );
}

// ── number_format deep ────────────────────────────────────────

#[test]
fn number_format_basic() {
    compile_ok(
        r#"<?php
echo number_format(1234567.891);
echo number_format(1234567.891, 2);
"#,
    );
}

#[test]
fn number_format_custom_separators() {
    compile_ok(
        r#"<?php
echo number_format(1234567.89, 2, ',', '.');  // European format
echo number_format(1234567.89, 2, '.', ' ');  // French format
"#,
    );
}

#[test]
fn number_format_no_decimals() {
    compile_ok(
        r#"<?php
echo number_format(9999999, 0, '.', ',');
"#,
    );
}

#[test]
fn number_format_small_numbers() {
    compile_ok(
        r#"<?php
echo number_format(0.005, 2);
echo number_format(0.0,   2);
echo number_format(-1.5,  1);
"#,
    );
}

// ── money_format alternative (intl NumberFormatter) ──────────

#[test]
fn intl_number_formatter_currency() {
    compile_ok(
        r#"<?php
if (class_exists('NumberFormatter')) {
    $fmt = new NumberFormatter('en_US', NumberFormatter::CURRENCY);
    echo $fmt->formatCurrency(1234.56, 'USD');
} else {
    echo '$1,234.56';
}
"#,
    );
}

// ── String padding ────────────────────────────────────────────

#[test]
fn str_pad_right() {
    compile_ok(
        r#"<?php
echo str_pad("hello", 10) . "|";
echo str_pad("hi", 8, "-") . "|";
"#,
    );
}

#[test]
fn str_pad_left() {
    compile_ok(
        r#"<?php
echo str_pad("42", 6, "0", STR_PAD_LEFT);
echo str_pad("x", 5, ".", STR_PAD_LEFT);
"#,
    );
}

#[test]
fn str_pad_both() {
    compile_ok(
        r#"<?php
echo str_pad("hi", 8, "-", STR_PAD_BOTH) . "|";
echo str_pad("a", 5, "*", STR_PAD_BOTH)  . "|";
"#,
    );
}

// ── wordwrap ──────────────────────────────────────────────────

#[test]
fn wordwrap_basic() {
    compile_ok(
        r#"<?php
$text = "The quick brown fox jumped over the lazy dog";
echo wordwrap($text, 15, "\n", false);
"#,
    );
}

#[test]
fn wordwrap_cut_long_words() {
    compile_ok(
        r#"<?php
$text = "A verylongwordthatcannotfit in normal wrapping";
echo wordwrap($text, 10, "\n", true);
"#,
    );
}

// ── chunk_split ───────────────────────────────────────────────

#[test]
fn chunk_split_basic() {
    compile_ok(
        r#"<?php
echo chunk_split("ABCDEFGH", 2, "-");
"#,
    );
}

#[test]
fn chunk_split_base64() {
    compile_ok(
        r#"<?php
$data = base64_encode(str_repeat("x", 60));
$formatted = chunk_split($data, 76, "\n");
echo strlen($formatted) > strlen($data) ? 'has newlines' : 'no newlines';
"#,
    );
}

// ── fprintf ───────────────────────────────────────────────────

#[test]
fn fprintf_to_stdout() {
    compile_ok(
        r#"<?php
$written = fprintf(STDOUT, "Value: %d\n", 42);
echo $written > 0 ? 'wrote' : 'nothing';
"#,
    );
}

// ── Practical sprintf patterns ────────────────────────────────

#[test]
fn sprintf_sql_like_pattern() {
    compile_ok(
        r#"<?php
// Template-based string building (illustrative, not real SQL)
function buildQuery(string $table, int $id): string {
    return sprintf("SELECT * FROM `%s` WHERE id = %d LIMIT 1", $table, $id);
}
echo buildQuery('users', 42);
"#,
    );
}

#[test]
fn sprintf_log_line() {
    compile_ok(
        r#"<?php
function logLine(string $level, string $msg, mixed ...$args): string {
    $formatted = empty($args) ? $msg : sprintf($msg, ...$args);
    return sprintf("[%s] %s", strtoupper($level), $formatted);
}
echo logLine('info', 'User %s logged in from %s', 'Alice', '127.0.0.1');
"#,
    );
}

#[test]
fn sprintf_currency_table() {
    compile_ok(
        r#"<?php
$items = [['Widget', 5, 9.99], ['Gadget', 2, 24.95], ['Doohickey', 1, 4.50]];
$total = 0.0;
foreach ($items as [$name, $qty, $price]) {
    $line = $qty * $price;
    $total += $line;
    printf("%-12s %3d @ %6.2f = %8.2f\n", $name, $qty, $price, $line);
}
printf("%s\n", str_repeat('-', 36));
printf("%-18s %16.2f\n", 'Total:', $total);
"#,
    );
}
