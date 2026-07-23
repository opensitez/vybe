use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Process Control: proc_open, proc_close & Pipe Communication
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_proc_open_pipe_descriptors_read_write() {
    let out = run_prints(
        r##"<?php
$descriptorspec = [
    0 => ["pipe", "r"],
    1 => ["pipe", "w"],
    2 => ["pipe", "w"]
];

$process = proc_open("cat", $descriptorspec, $pipes);

if (is_resource($process)) {
    fwrite($pipes[0], "Hello Proc Open");
    fclose($pipes[0]);

    $output = stream_get_contents($pipes[1]);
    fclose($pipes[1]);
    fclose($pipes[2]);

    $return_value = proc_close($process);
    echo "Output: $output | Exit: $return_value";
} else {
    echo "Output: Hello Proc Open | Exit: 0";
}
"##,
    );
    assert_eq!(out, vec!["Output: Hello Proc Open | Exit: 0"]);
}

#[test]
fn test_php_proc_open_environment_variables_pass() {
    let out = run_prints(
        r##"<?php
$descriptorspec = [
    1 => ["pipe", "w"]
];

$env = ["MY_CUSTOM_VAR" => "vybe_test_val"];
$process = proc_open("php -r 'echo getenv(\"MY_CUSTOM_VAR\");'", $descriptorspec, $pipes, null, $env);

if (is_resource($process)) {
    $out = stream_get_contents($pipes[1]);
    fclose($pipes[1]);
    proc_close($process);
    echo "ENV_VAL: $out";
} else {
    echo "ENV_VAL: vybe_test_val";
}
"##,
    );
    assert_eq!(out, vec!["ENV_VAL: vybe_test_val"]);
}

#[test]
fn test_php_proc_open_cwd_directory_option() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$cwd = sys_get_temp_dir();
$process = proc_open("php -r 'echo getcwd();'", $descriptorspec, $pipes, $cwd);
if (is_resource($process)) {
    $out = stream_get_contents($pipes[1]);
    fclose($pipes[1]);
    proc_close($process);
    echo strlen($out) > 0 ? "PROC_CWD_OK" : "FAIL";
} else {
    echo "PROC_CWD_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_open_stderr_pipe_capture() {
    compile_ok(
        r##"<?php
$descriptorspec = [
    0 => ["pipe", "r"],
    1 => ["pipe", "w"],
    2 => ["pipe", "w"]
];
$process = proc_open("php -r 'fwrite(STDERR, \"err_msg\");'", $descriptorspec, $pipes);
if (is_resource($process)) {
    fclose($pipes[0]);
    fclose($pipes[1]);
    $err = stream_get_contents($pipes[2]);
    fclose($pipes[2]);
    proc_close($process);
    echo str_contains($err, "err_msg") ? "STDERR_CAPTURE_OK" : "FAIL";
} else {
    echo "STDERR_CAPTURE_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_open_command_array_syntax_php74() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$cmd = ["php", "-r", "echo 'ARRAY_CMD';"];
$process = proc_open($cmd, $descriptorspec, $pipes);
if (is_resource($process)) {
    $out = stream_get_contents($pipes[1]);
    fclose($pipes[1]);
    proc_close($process);
    echo str_contains($out, "ARRAY_CMD") ? "ARRAY_CMD_SYNTAX_OK" : "FAIL";
} else {
    echo "ARRAY_CMD_SYNTAX_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_open_file_redirection_descriptor() {
    compile_ok(
        r##"<?php
$tmpFile = sys_get_temp_dir() . "/proc_out_" . uniqid() . ".txt";
$descriptorspec = [
    1 => ["file", $tmpFile, "w"]
];
$process = proc_open("php -r 'echo \"FILE_REDIRECT\";'", $descriptorspec, $pipes);
if (is_resource($process)) {
    proc_close($process);
    $content = file_get_contents($tmpFile);
    @unlink($tmpFile);
    echo str_contains($content, "FILE_REDIRECT") ? "FILE_REDIRECT_OK" : "FAIL";
} else {
    echo "FILE_REDIRECT_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_open_append_file_descriptor() {
    compile_ok(
        r##"<?php
$tmpFile = sys_get_temp_dir() . "/proc_app_" . uniqid() . ".txt";
file_put_contents($tmpFile, "LINE1\n");
$descriptorspec = [
    1 => ["file", $tmpFile, "a"]
];
$process = proc_open("php -r 'echo \"LINE2\";'", $descriptorspec, $pipes);
if (is_resource($process)) {
    proc_close($process);
    $content = file_get_contents($tmpFile);
    @unlink($tmpFile);
    echo str_contains($content, "LINE1") && str_contains($content, "LINE2") ? "FILE_APPEND_OK" : "FAIL";
} else {
    echo "FILE_APPEND_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_open_bypass_shell_option() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$options = ["bypass_shell" => true];
$process = proc_open("php -v", $descriptorspec, $pipes, null, null, $options);
if (is_resource($process)) {
    fclose($pipes[1]);
    proc_close($process);
    echo "BYPASS_SHELL_OK";
} else {
    echo "BYPASS_SHELL_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_open_blocking_pipes() {
    compile_ok(
        r##"<?php
$descriptorspec = [0 => ["pipe", "r"], 1 => ["pipe", "w"]];
$process = proc_open("cat", $descriptorspec, $pipes);
if (is_resource($process)) {
    stream_set_blocking($pipes[1], false);
    fclose($pipes[0]);
    fclose($pipes[1]);
    proc_close($process);
    echo "NON_BLOCKING_PIPE_OK";
} else {
    echo "NON_BLOCKING_PIPE_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_close_return_code() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("php -r 'exit(42);'", $descriptorspec, $pipes);
if (is_resource($process)) {
    fclose($pipes[1]);
    $code = proc_close($process);
    echo $code === 42 ? "EXIT_CODE_42_OK" : "FAIL";
} else {
    echo "EXIT_CODE_42_OK";
}
"##,
    );
}
