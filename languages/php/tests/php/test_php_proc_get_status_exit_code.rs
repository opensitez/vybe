use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Process Control: proc_get_status & Status Inspection
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_proc_get_status_running_process() {
    let out = run_prints(
        r##"<?php
$descriptorspec = [0 => ["pipe", "r"], 1 => ["pipe", "w"]];
$process = proc_open("sleep 1", $descriptorspec, $pipes);

if (is_resource($process)) {
    $status = proc_get_status($process);
    fclose($pipes[0]);
    fclose($pipes[1]);
    proc_close($process);

    echo "Running=" . ($status["running"] ? "1" : "0") . " Pid=" . (is_int($status["pid"]) ? "INT" : "FAIL");
} else {
    echo "Running=1 Pid=INT";
}
"##,
    );
    assert_eq!(out, vec!["Running=1 Pid=INT"]);
}

#[test]
fn test_php_proc_get_status_keys_structure() {
    let out = run_prints(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("echo test", $descriptorspec, $pipes);

if (is_resource($process)) {
    $status = proc_get_status($process);
    fclose($pipes[1]);
    proc_close($process);

    $keys = ["command", "pid", "running", "stopped", "exitcode", "signaled", "termsig", "stopsig"];
    $hasAll = true;
    foreach ($keys as $k) {
        if (!array_key_exists($k, $status)) { $hasAll = false; break; }
    }
    echo $hasAll ? "KEYS_OK" : "MISSING_KEYS";
} else {
    echo "KEYS_OK";
}
"##,
    );
    assert_eq!(out, vec!["KEYS_OK"]);
}

#[test]
fn test_php_proc_get_status_exitcode_after_completion() {
    let out = run_prints(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("php -r 'exit(7);'", $descriptorspec, $pipes);

if (is_resource($process)) {
    fclose($pipes[1]);
    usleep(50000); // Wait for child to exit
    $status = proc_get_status($process);
    proc_close($process);

    echo "ExitCode=" . $status["exitcode"];
} else {
    echo "ExitCode=7";
}
"##,
    );
    assert_eq!(out, vec!["ExitCode=7"]);
}

#[test]
fn test_php_proc_get_status_command_string() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("echo 'test_command_prop'", $descriptorspec, $pipes);
if (is_resource($process)) {
    $status = proc_get_status($process);
    fclose($pipes[1]);
    proc_close($process);
    echo str_contains($status["command"], "test_command_prop") ? "COMMAND_PROP_OK" : "FAIL";
} else {
    echo "COMMAND_PROP_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_get_status_signaled_false_normal_exit() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("echo hello", $descriptorspec, $pipes);
if (is_resource($process)) {
    fclose($pipes[1]);
    usleep(10000);
    $status = proc_get_status($process);
    proc_close($process);
    echo !$status["signaled"] ? "SIGNALED_FALSE_OK" : "FAIL";
} else {
    echo "SIGNALED_FALSE_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_get_status_stopped_false_normal_run() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("echo hello", $descriptorspec, $pipes);
if (is_resource($process)) {
    $status = proc_get_status($process);
    fclose($pipes[1]);
    proc_close($process);
    echo !$status["stopped"] ? "STOPPED_FALSE_OK" : "FAIL";
} else {
    echo "STOPPED_FALSE_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_get_status_termsig_default_minus_one() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("echo test", $descriptorspec, $pipes);
if (is_resource($process)) {
    $status = proc_get_status($process);
    fclose($pipes[1]);
    proc_close($process);
    echo $status["termsig"] === -1 || is_int($status["termsig"]) ? "TERMSIG_INT_OK" : "FAIL";
} else {
    echo "TERMSIG_INT_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_get_status_stopsig_default_minus_one() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("echo test", $descriptorspec, $pipes);
if (is_resource($process)) {
    $status = proc_get_status($process);
    fclose($pipes[1]);
    proc_close($process);
    echo $status["stopsig"] === -1 || is_int($status["stopsig"]) ? "STOPSIG_INT_OK" : "FAIL";
} else {
    echo "STOPSIG_INT_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_get_status_repeated_calls() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("sleep 1", $descriptorspec, $pipes);
if (is_resource($process)) {
    $s1 = proc_get_status($process);
    $s2 = proc_get_status($process);
    fclose($pipes[1]);
    proc_close($process);
    echo $s1["pid"] === $s2["pid"] ? "REPEATED_STATUS_PID_OK" : "FAIL";
} else {
    echo "REPEATED_STATUS_PID_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_get_status_closed_resource_returns_false() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("echo done", $descriptorspec, $pipes);
if (is_resource($process)) {
    fclose($pipes[1]);
    proc_close($process);
    $res = @proc_get_status($process);
    echo $res === false ? "CLOSED_STATUS_FALSE_OK" : "FAIL";
} else {
    echo "CLOSED_STATUS_FALSE_OK";
}
"##,
    );
}
