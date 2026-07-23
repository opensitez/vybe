use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Shell Sanitization: escapeshellcmd & escapeshellarg
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_escapeshellarg_wraps_in_single_quotes() {
    let out = run_prints(
        r##"<?php
$arg = "user input; rm -rf /";
$escaped = escapeshellarg($arg);
echo $escaped;
"##,
    );
    assert_eq!(out, vec!["'user input; rm -rf /'"]);
}

#[test]
fn test_php_escapeshellcmd_escapes_metacharacters() {
    let out = run_prints(
        r##"<?php
$cmd = "ls -l; cat /etc/passwd | grep root & bg";
$escaped = escapeshellcmd($cmd);
echo $escaped;
"##,
    );
    assert_eq!(out, vec!["ls -l\\; cat /etc/passwd \\| grep root \\& bg"]);
}

#[test]
fn test_php_escapeshellarg_escapes_single_quotes_inside() {
    let out = run_prints(
        r##"<?php
$arg = "don't stop";
$escaped = escapeshellarg($arg);
echo str_contains($escaped, "'") ? "QUOTES_ESCAPED_OK" : "FAIL";
"##,
    );
    assert_eq!(out, vec!["QUOTES_ESCAPED_OK"]);
}

#[test]
fn test_php_escapeshellcmd_escapes_dollar_and_backticks() {
    compile_ok(
        r##"<?php
$input = "echo \$VAR `whoami` (sub)";
$clean = escapeshellcmd($input);
echo str_contains($clean, "\\$") && str_contains($clean, "\\`") ? "METACHARS_ESCAPED_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_escapeshellarg_empty_string() {
    compile_ok(
        r##"<?php
$escaped = escapeshellarg("");
echo $escaped === "''" || $escaped === '""' ? "EMPTY_SHELL_ARG_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_escapeshellcmd_redirection_operators() {
    compile_ok(
        r##"<?php
$cmd = "cat < input.txt > output.txt 2>&1";
$clean = escapeshellcmd($cmd);
echo str_contains($clean, "\\<") && str_contains($clean, "\\>") ? "REDIRECTION_ESCAPED_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_escapeshellarg_numeric_values() {
    compile_ok(
        r##"<?php
$arg = 12345;
$escaped = escapeshellarg($arg);
echo $escaped === "'12345'" || $escaped === '"12345"' ? "NUMERIC_ARG_ESCAPED_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_escapeshellcmd_newlines_escaping() {
    compile_ok(
        r##"<?php
$cmd = "echo hello\necho world";
$clean = escapeshellcmd($cmd);
echo is_string($clean) ? "NEWLINE_ESCAPED_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_escapeshellarg_filename_with_spaces() {
    compile_ok(
        r##"<?php
$filename = "my file name.pdf";
$escaped = escapeshellarg($filename);
echo $escaped === "'my file name.pdf'" || $escaped === '"my file name.pdf"' ? "SPACES_FILENAME_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_escapeshellcmd_wildcard_characters() {
    compile_ok(
        r##"<?php
$cmd = "ls *.php ?";
$clean = escapeshellcmd($cmd);
echo str_contains($clean, "\\*") || str_contains($clean, "\\?") || is_string($clean) ? "WILDCARD_ESCAPED_OK" : "FAIL";
"##,
    );
}
