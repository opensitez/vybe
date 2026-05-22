use super::helpers::compile_ok;

// ── printf with multiple format specifiers ────────────────────

#[test] fn printf_multiple_specifiers() {
    compile_ok(r#"<?php
$written = printf("Name: %s, Age: %d, Score: %.2f\n", "Alice", 30, 98.5);
echo $written > 0 ? 'wrote bytes' : 'nothing written';
"#);
}

// ── fprintf to STDOUT ─────────────────────────────────────────

#[test] fn fprintf_to_stdout_explicit() {
    compile_ok(r#"<?php
$written = fprintf(STDOUT, "Item: %s costs $%.2f\n", "widget", 4.99);
echo $written > 0 ? 'wrote' : 'nothing';
"#);
}

// ── sprintf with padding and alignment ───────────────────────

#[test] fn sprintf_padding_and_alignment() {
    compile_ok(r#"<?php
$left  = sprintf('%-10s|', 'hi');
$right = sprintf('%010d',  42);
echo $left;
echo $right;
"#);
}

// ── sprintf with %+ sign flag ─────────────────────────────────

#[test] fn sprintf_sign_flag() {
    compile_ok(r#"<?php
echo sprintf('%+d', 42);
echo sprintf('%+d', -42);
echo sprintf('%+.2f', 3.14);
echo sprintf('%+.2f', -3.14);
"#);
}

// ── sprintf with %x %o %b ────────────────────────────────────

#[test] fn sprintf_hex_octal_binary() {
    compile_ok(r#"<?php
echo sprintf('%x', 255);
echo sprintf('%X', 255);
echo sprintf('%o', 8);
echo sprintf('%b', 10);
echo sprintf('%08b', 10);
"#);
}

// ── sprintf with %e scientific notation ──────────────────────

#[test] fn sprintf_scientific_notation() {
    compile_ok(r#"<?php
echo sprintf('%e', 123456.789);
echo sprintf('%E', 0.000123);
echo sprintf('%.3e', 9876.5432);
"#);
}

// ── vprintf — printf with array ───────────────────────────────

#[test] fn vprintf_with_array() {
    compile_ok(r#"<?php
$args = ['Charlie', 7, 99.9];
$written = vprintf("Player: %s, Level: %d, HP: %.1f\n", $args);
echo $written > 0 ? 'ok' : 'fail';
"#);
}

// ── vsprintf — sprintf with array ────────────────────────────

#[test] fn vsprintf_with_array() {
    compile_ok(r#"<?php
$args = ['PHP', '8.3', 'Stable'];
$result = vsprintf('%s %s (%s)', $args);
echo $result;
"#);
}

// ── print_r with return=true ──────────────────────────────────

#[test] fn print_r_return_true() {
    compile_ok(r#"<?php
$data = ['a' => 1, 'b' => [2, 3]];
$output = print_r($data, true);
echo is_string($output) ? 'string' : 'not string';
echo strlen($output) > 0 ? ':non-empty' : ':empty';
"#);
}

// ── var_export with return=true ───────────────────────────────

#[test] fn var_export_return_true() {
    compile_ok(r#"<?php
$data = ['x' => 1, 'y' => 'hello', 'z' => true];
$code = var_export($data, true);
echo is_string($code) ? 'string' : 'not string';
echo strlen($code) > 0 ? ':non-empty' : ':empty';
"#);
}

// ── var_dump of nested structure ──────────────────────────────

#[test] fn var_dump_nested() {
    compile_ok(r#"<?php
$obj = new stdClass();
$obj->name = 'test';
$obj->values = [1, 2, 3];
ob_start();
var_dump($obj);
$output = ob_get_clean();
echo strlen($output) > 0 ? 'dumped' : 'empty';
"#);
}

// ── ob_start + ob_get_clean ───────────────────────────────────

#[test] fn ob_start_get_clean_capture() {
    compile_ok(r#"<?php
ob_start();
echo 'captured output';
$content = ob_get_clean();
echo 'got: ' . $content;
"#);
}

// ── ob_start + ob_end_flush ───────────────────────────────────

#[test] fn ob_start_end_flush_sends() {
    compile_ok(r#"<?php
ob_start();
echo 'flushed content';
ob_end_flush();
"#);
}

// ── ob_get_contents without flushing ─────────────────────────

#[test] fn ob_get_contents_no_flush() {
    compile_ok(r#"<?php
ob_start();
echo 'peek';
$buf = ob_get_contents();
ob_end_clean();
echo 'peeked: ' . $buf;
"#);
}

// ── ob_get_level nesting level ────────────────────────────────

#[test] fn ob_get_level_nesting() {
    compile_ok(r#"<?php
$l0 = ob_get_level();
ob_start();
$l1 = ob_get_level();
ob_start();
$l2 = ob_get_level();
ob_end_clean();
ob_end_clean();
$l3 = ob_get_level();
echo "$l0,$l1,$l2,$l3";
"#);
}
