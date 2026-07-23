use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Output Buffering — ob_start, ob_get_clean, ob_get_contents, ob_end_clean, ob_flush, nested buffering
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_ob_start_ob_get_clean_capture() {
    let out = run_prints(
        r#"<?php
ob_start();
echo "Buffered HTML Content";
$captured = ob_get_clean();
echo "Captured: $captured";
"#,
    );
    assert_eq!(out, vec!["Captured: Buffered HTML Content"]);
}

#[test]
fn test_php_ob_get_contents_and_end_clean() {
    let out = run_prints(
        r#"<?php
ob_start();
echo "Internal output";
$contents = ob_get_contents();
ob_end_clean();
echo "Retrieved: $contents";
"#,
    );
    assert_eq!(out, vec!["Retrieved: Internal output"]);
}

#[test]
fn test_php_nested_output_buffering_levels() {
    let out = run_prints(
        r#"<?php
ob_start(); // Level 1
echo "Level 1";
ob_start(); // Level 2
echo "Level 2";
$l2 = ob_get_clean();
$l1 = ob_get_clean();
echo "L1=$l1 | L2=$l2";
"#,
    );
    assert_eq!(out, vec!["L1=Level 1 | L2=Level 2"]);
}

#[test]
fn test_php_ob_start_callback_processor() {
    let out = run_prints(
        r#"<?php
ob_start(function($buffer) {
    return strtoupper($buffer);
});
echo "lowercase text";
ob_end_flush();
"#,
    );
    assert_eq!(out, vec!["LOWERCASE TEXT"]);
}

#[test]
fn test_php_ob_get_level_and_ob_get_status() {
    compile_ok(
        r#"<?php
$initialLevel = ob_get_level();
ob_start();
echo "Level active: " . (ob_get_level() - $initialLevel);
$status = ob_get_status();
print_r($status);
ob_end_clean();
"#,
    );
}

#[test]
fn test_php_ob_clean_clears_buffer_without_closing() {
    compile_ok(
        r#"<?php
ob_start();
echo "discarded text";
ob_clean();
echo "kept text";
$output = ob_get_clean();
echo $output;
"#,
    );
}

#[test]
fn test_php_ob_implicit_flush_setting() {
    compile_ok(
        r#"<?php
ob_implicit_flush(true);
echo "immediate output";
ob_implicit_flush(false);
"#,
    );
}

#[test]
fn test_php_ob_flush_sends_buffer_to_output() {
    compile_ok(
        r#"<?php
ob_start();
echo "flushed part 1\n";
ob_flush();
echo "flushed part 2\n";
ob_end_clean();
"#,
    );
}

#[test]
fn test_php_ob_gzhandler_compression_check() {
    compile_ok(
        r#"<?php
if (extension_loaded('zlib')) {
    ob_start('ob_gzhandler');
    echo "Compressed page content";
    ob_end_clean();
}
"#,
    );
}

#[test]
fn test_php_output_buffering_template_include_capture() {
    compile_ok(
        r#"<?php
function renderTemplate(string $content): string {
    ob_start();
    echo "<div>$content</div>";
    return ob_get_clean();
}

echo renderTemplate("Hello World");
"#,
    );
}
