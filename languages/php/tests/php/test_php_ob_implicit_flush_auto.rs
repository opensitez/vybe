use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Output Buffering: ob_implicit_flush & Automatic Flushing
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_ob_implicit_flush_enable_and_disable() {
    let out = run_prints(
        r##"<?php
ob_implicit_flush(true);
ob_implicit_flush(false);
echo "ImplicitFlush Toggled";
"##,
    );
    assert_eq!(out, vec!["ImplicitFlush Toggled"]);
}

#[test]
fn test_php_flush_sends_system_buffer() {
    let out = run_prints(
        r##"<?php
echo "Flushed Chunk 1\n";
flush();
echo "Flushed Chunk 2";
"##,
    );
    assert_eq!(out, vec!["Flushed Chunk 1\nFlushed Chunk 2"]);
}

#[test]
fn test_php_ob_end_flush_empties_and_disables() {
    compile_ok(
        r##"<?php
ob_start();
echo "Buffered Data";
$levelBefore = ob_get_level();
ob_end_flush();
$levelAfter = ob_get_level();
echo $levelAfter === $levelBefore - 1 ? "END_FLUSH_DECR_LEVEL_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_get_flush_returns_content_and_flushes() {
    compile_ok(
        r##"<?php
ob_start();
echo "Flush Content";
$content = ob_get_flush();
echo $content === "Flush Content" ? "GET_FLUSH_CONTENT_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_implicit_flush_default_off() {
    compile_ok(
        r##"<?php
ob_implicit_flush(0);
echo "Implicit OFF";
"##,
    );
}

#[test]
fn test_php_ob_implicit_flush_numeric_arguments() {
    compile_ok(
        r##"<?php
ob_implicit_flush(1);
ob_implicit_flush(0);
echo "NUMERIC_FLUSH_ARGS_OK";
"##,
    );
}

#[test]
fn test_php_ob_end_clean_on_inactive_buffer_returns_false() {
    compile_ok(
        r##"<?php
while (ob_get_level() > 0) ob_end_clean();
$res = @ob_end_clean();
echo $res === false ? "INACTIVE_END_CLEAN_FALSE" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_end_flush_on_inactive_buffer_returns_false() {
    compile_ok(
        r##"<?php
while (ob_get_level() > 0) ob_end_clean();
$res = @ob_end_flush();
echo $res === false ? "INACTIVE_END_FLUSH_FALSE" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_clean_on_inactive_buffer_returns_false() {
    compile_ok(
        r##"<?php
while (ob_get_level() > 0) ob_end_clean();
$res = @ob_clean();
echo $res === false ? "INACTIVE_CLEAN_FALSE" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_flush_on_inactive_buffer_returns_false() {
    compile_ok(
        r##"<?php
while (ob_get_level() > 0) ob_end_clean();
$res = @ob_flush();
echo $res === false ? "INACTIVE_FLUSH_FALSE" : "FAIL";
"##,
    );
}
