use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Output Buffering Nested Handlers — ob_start nested handlers, ob_get_length, ob_get_level, ob_list_handlers, ob_end_flush
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_ob_list_handlers_active_stack() {
    let out = run_prints(
        r#"<?php
ob_start();
ob_start();
$handlers = ob_list_handlers();
echo implode(", ", $handlers);
ob_end_clean();
ob_end_clean();
"#,
    );
    assert_eq!(out, vec!["default output handler, default output handler"]);
}

#[test]
fn test_php_ob_get_length_buffer_byte_count() {
    let out = run_prints(
        r#"<?php
ob_start();
echo "1234567890";
$len = ob_get_length();
ob_end_clean();
echo "Length: $len";
"#,
    );
    assert_eq!(out, vec!["Length: 10"]);
}

#[test]
fn test_php_ob_end_flush_emits_to_outer_buffer() {
    let out = run_prints(
        r#"<?php
ob_start(); // Outer
echo "OUTER_START ";
ob_start(); // Inner
echo "INNER_TEXT ";
ob_end_flush(); // Flushes Inner into Outer
echo "OUTER_END";
$final = ob_get_clean();
echo $final;
"#,
    );
    assert_eq!(out, vec!["OUTER_START INNER_TEXT OUTER_END"]);
}

#[test]
fn test_php_ob_start_chunk_size_auto_flush() {
    compile_ok(
        r#"<?php
$flushedChunks = [];
ob_start(function($buffer) use (&$flushedChunks) {
    $flushedChunks[] = $buffer;
    return $buffer;
}, chunk_size: 10);

echo "1234567890"; // Should trigger chunk flush
echo "abcdefghij";
ob_end_clean();
"#,
    );
}

#[test]
fn test_php_ob_start_flags_erase_write_flush() {
    compile_ok(
        r#"<?php
// Buffer that cannot be cleaned, only flushed
ob_start(flags: PHP_OUTPUT_HANDLER_FLUSHALBLE | PHP_OUTPUT_HANDLER_REMOVABLE);
echo "non_erasable_content";
ob_end_flush();
"#,
    );
}

#[test]
fn test_php_ob_get_status_full_stack_details() {
    compile_ok(
        r#"<?php
ob_start();
$statuses = ob_get_status(full_status: true);
echo is_array($statuses) && count($statuses) > 0 ? "STATUS_OK" : "FAIL";
ob_end_clean();
"#,
    );
}

#[test]
fn test_php_flush_system_output_buffer() {
    compile_ok(
        r#"<?php
echo "Flushing output stream...";
flush();
"#,
    );
}

#[test]
fn test_php_ob_clean_and_refill() {
    compile_ok(
        r#"<?php
ob_start();
echo "wrong data";
ob_clean();
echo "correct data";
$res = ob_get_clean();
echo $res;
"#,
    );
}

#[test]
fn test_php_ob_start_multiple_custom_filters() {
    compile_ok(
        r#"<?php
ob_start(fn($s) => str_replace("foo", "bar", $s));
ob_start(fn($s) => strtoupper($s));
echo "foo text";
ob_end_flush();
ob_end_flush();
"#,
    );
}

#[test]
fn test_php_ob_handler_constants() {
    compile_ok(
        r#"<?php
echo PHP_OUTPUT_HANDLER_START . " " . PHP_OUTPUT_HANDLER_WRITE . " " . PHP_OUTPUT_HANDLER_END;
"#,
    );
}
