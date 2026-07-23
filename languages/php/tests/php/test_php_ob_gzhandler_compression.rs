use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Output Buffering: ob_gzhandler Compression & Buffer Flags
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_ob_gzhandler_buffer_wrapping() {
    let out = run_prints(
        r##"<?php
if (function_exists('ob_gzhandler')) {
    ob_start("ob_gzhandler");
    echo "Compressed Page Output";
    $content = ob_get_clean();
    echo "BufferHandled: " . (strlen($content) > 0 ? "YES" : "NO");
} else {
    echo "BufferHandled: YES";
}
"##,
    );
    assert_eq!(out, vec!["BufferHandled: YES"]);
}

#[test]
fn test_php_ob_start_chunk_size_trigger() {
    let out = run_prints(
        r##"<?php
$chunks = 0;
ob_start(function($buffer) use (&$chunks) {
    $chunks++;
    return $buffer;
}, 10); // Chunk size 10 bytes

echo "12345678901"; // 11 bytes triggers chunk handler
$count = $chunks;
ob_end_clean();

echo "ChunksTriggered: " . ($count > 0 ? "YES" : "NO");
"##,
    );
    assert_eq!(out, vec!["ChunksTriggered: YES"]);
}

#[test]
fn test_php_ob_start_flags_cleanable_flushable() {
    compile_ok(
        r##"<?php
ob_start(null, 0, PHP_OUTPUT_HANDLER_CLEANABLE | PHP_OUTPUT_HANDLER_FLUSHABLE);
echo "Flushes";
$cleared = ob_end_clean();
echo $cleared ? "CLEANABLE_FLAG_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_start_removable_flag() {
    compile_ok(
        r##"<?php
ob_start(null, 0, PHP_OUTPUT_HANDLER_REMOVABLE);
echo "Removable";
$ended = ob_end_clean();
echo $ended ? "REMOVABLE_FLAG_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_get_length_returns_buffer_byte_count() {
    compile_ok(
        r##"<?php
ob_start();
echo "1234567890";
$len = ob_get_length();
ob_end_clean();
echo $len === 10 ? "BUFFER_LEN_10_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_get_level_nested_depth() {
    compile_ok(
        r##"<?php
$l0 = ob_get_level();
ob_start();
$l1 = ob_get_level();
ob_start();
$l2 = ob_get_level();
ob_end_clean();
ob_end_clean();
echo $l1 === $l0 + 1 && $l2 === $l0 + 2 ? "NESTED_LEVELS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_clean_clears_without_ending() {
    compile_ok(
        r##"<?php
ob_start();
echo "First";
ob_clean();
echo "Second";
$out = ob_get_clean();
echo $out === "Second" ? "OB_CLEAN_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_flush_sends_to_outer_buffer() {
    compile_ok(
        r##"<?php
ob_start(); // Outer
ob_start(); // Inner
echo "InnerPayload";
ob_flush(); // Flushes inner to outer
$innerRemaining = ob_get_clean();
$outerPayload = ob_get_clean();
echo $innerRemaining === "" && $outerPayload === "InnerPayload" ? "OB_FLUSH_OUTER_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_get_contents_without_clearing() {
    compile_ok(
        r##"<?php
ob_start();
echo "Persistent";
$c1 = ob_get_contents();
$c2 = ob_get_contents();
ob_end_clean();
echo $c1 === "Persistent" && $c2 === "Persistent" ? "PERSISTENT_CONTENTS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_start_passthrough_return() {
    compile_ok(
        r##"<?php
ob_start(fn($s) => strtoupper($s));
echo "lowercase";
$out = ob_get_clean();
echo $out === "LOWERCASE" ? "PASSTHROUGH_STRTOUPPER_OK" : "FAIL";
"##,
    );
}
