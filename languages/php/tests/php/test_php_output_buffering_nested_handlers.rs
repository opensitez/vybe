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
ob_end_clean();
ob_end_clean();
echo implode(", ", $handlers);
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

#[test]
fn test_php_ob_get_status_runtime() {
    let out = run_prints(
        r#"<?php
ob_start();
echo "abc";
$status = ob_get_status();
ob_end_clean();
echo $status['name'] . '|' . $status['level'];
"#,
    );
    assert_eq!(out, vec!["default output handler|0"]);
}

#[test]
fn test_php_ob_get_status_full_true_runtime() {
    let out = run_prints(
        r#"<?php
ob_start();
echo "x";
$status = ob_get_status(true);
ob_end_clean();
echo is_array($status) ? (is_array(array_values($status)[0]) ? 'full' : 'partial') : 'false';
"#,
    );
    assert_eq!(out, vec!["full"]);
}

#[test]
fn test_php_ob_start_with_nested_transform_chain_runtime() {
    let out = run_prints(
        r#"<?php
ob_start(function(string $chunk): string {
    return '[' . $chunk . ']';
});
ob_start(function(string $chunk): string {
    return strtoupper($chunk);
});
echo 'ok';
$inner = ob_get_clean();
echo $inner;
ob_end_flush();
"#,
    );
    assert_eq!(out, vec!["[OK]"]);
}

#[test]
fn test_php_ob_start_with_chunk_size_runtime() {
    let out = run_prints(
        r#"<?php
ob_start(null, 32, false);
echo str_repeat('a', 5);
$inside = ob_get_contents();
ob_end_clean();
echo $inside . '|';
"#,
    );
    assert_eq!(out, vec!["aaaaa|"]);
}

#[test]
fn test_php_ob_get_length_after_nested_clean_runtime() {
    let out = run_prints(
        r#"<?php
ob_start();
ob_start();
echo 'inner';
$inner_len = ob_get_length();
ob_end_clean();
$outer_len = ob_get_length();
ob_end_clean();
echo $inner_len . '|' . $outer_len;
"#,
    );
    assert_eq!(out, vec!["5|0"]);
}

#[test]
fn test_php_ob_end_clean_pops_only_top_handler_runtime() {
    let out = run_prints(
        r#"<?php
ob_start();
ob_start();
echo 'inner';
ob_end_clean();
echo 'outer';
$status = ob_get_status(true);
$n = is_array($status) ? count($status) : -1;
$contents = ob_get_contents();
ob_end_clean();
echo $contents . '|' . $n . '|' . $contents;
"#,
    );
    assert_eq!(out, vec!["outer|1|outer"]);
}

#[test]
fn test_php_ob_get_level_for_nested_buffers_runtime() {
    let out = run_prints(
        r#"<?php
ob_start();
$l1 = ob_get_level();
ob_start();
$l2 = ob_get_level();
ob_end_clean();
$l3 = ob_get_level();
ob_end_clean();
echo $l1 . '|' . $l2 . '|' . $l3;
"#,
    );
    assert_eq!(out, vec!["1|2|1"]);
}

#[test]
fn test_php_ob_start_handler_with_returning_empty_discard_runtime() {
    let out = run_prints(
        r#"<?php
ob_start(function(string $chunk): string {
    return '';
});
echo 'should disappear';
ob_end_flush();
echo 'after';
"#,
    );
    assert_eq!(out, vec!["after"]);
}

#[test]
fn test_php_ob_start_with_callback_receives_full_chunk_runtime() {
    let out = run_prints(
        r#"<?php
ob_start(function(string $chunk): string {
    return strtoupper($chunk);
});
echo 'a';
echo 'b';
$c = ob_get_clean();
echo $c;
"#,
    );
    assert_eq!(out, vec!["AB"]);
}
