use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: CLI Arguments & Environment — $argc, $argv, getopt(), getenv(), putenv(), php_sapi_name(), sys_get_temp_dir()
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_getopt_short_and_long_options() {
    let out = run_prints(
        r#"<?php
// Simulate CLI args: -f value --required=10
$_SERVER['argv'] = ['script.php', '-f', 'bar', '--required=10'];
$_SERVER['argc'] = 4;

$options = getopt("f:", ["required:"]);
echo "f=" . $options["f"] . " required=" . $options["required"];
"#,
    );
    assert_eq!(out, vec!["f=bar required=10"]);
}

#[test]
fn test_php_getenv_putenv_environment_mutation() {
    let out = run_prints(
        r#"<?php
putenv("VYBE_ENV_VAR=active_test_123");
echo getenv("VYBE_ENV_VAR");
"#,
    );
    assert_eq!(out, vec!["active_test_123"]);
}

#[test]
fn test_php_php_sapi_name_cli_detection() {
    let out = run_prints(
        r#"<?php
$sapi = php_sapi_name();
echo (strlen($sapi) > 0) ? "SAPI_AVAILABLE" : "NO_SAPI";
"#,
    );
    assert_eq!(out, vec!["SAPI_AVAILABLE"]);
}

#[test]
fn test_php_sys_get_temp_dir_directory_path() {
    let out = run_prints(
        r#"<?php
$tmp = sys_get_temp_dir();
echo is_dir($tmp) ? "TEMP_DIR_EXISTS" : "TEMP_DIR_MISSING";
"#,
    );
    assert_eq!(out, vec!["TEMP_DIR_EXISTS"]);
}

#[test]
fn test_php_argc_argv_global_variables() {
    compile_ok(
        r#"<?php
global $argc, $argv;
echo "Args count: " . (is_array($argv) ? count($argv) : 0);
"#,
    );
}

#[test]
fn test_php_getopt_flag_without_values() {
    compile_ok(
        r#"<?php
$_SERVER['argv'] = ['app', '-v', '-q', '--debug'];
$opts = getopt("vq", ["debug"]);
echo isset($opts["v"]) && isset($opts["debug"]) ? "FLAGS_OK" : "FLAGS_FAIL";
"#,
    );
}

#[test]
fn test_php_get_current_user_and_process_id() {
    compile_ok(
        r#"<?php
$pid = getmypid();
$user = get_current_user();
echo "PID=$pid USER=$user";
"#,
    );
}

#[test]
fn test_php_cli_set_process_title() {
    compile_ok(
        r#"<?php
if (function_exists('cli_set_process_title')) {
    @cli_set_process_title("vybe-worker");
    echo cli_get_process_title();
}
"#,
    );
}

#[test]
fn test_php_memory_get_usage_and_peak() {
    compile_ok(
        r#"<?php
$alloc = memory_get_usage();
$peak = memory_get_peak_usage();
echo ($alloc > 0 && $peak >= $alloc) ? "MEM_USAGE_OK" : "MEM_FAIL";
"#,
    );
}

#[test]
fn test_php_getenv_local_array_parsing() {
    compile_ok(
        r#"<?php
$env = getenv();
echo is_array($env) ? "ENV_ARRAY" : "FAIL";
"#,
    );
}
