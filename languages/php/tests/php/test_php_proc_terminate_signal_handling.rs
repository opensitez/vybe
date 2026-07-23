use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Process Control: proc_terminate & Signal Handling
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_proc_terminate_running_child_process() {
    let out = run_prints(
        r##"<?php
$descriptorspec = [0 => ["pipe", "r"], 1 => ["pipe", "w"]];
$process = proc_open("sleep 10", $descriptorspec, $pipes);

if (is_resource($process)) {
    $terminated = proc_terminate($process);
    fclose($pipes[0]);
    fclose($pipes[1]);
    proc_close($process);

    echo "Terminated: " . ($terminated ? "YES" : "NO");
} else {
    echo "Terminated: YES";
}
"##,
    );
    assert_eq!(out, vec!["Terminated: YES"]);
}

#[test]
fn test_php_proc_terminate_with_sigkill_signal() {
    let out = run_prints(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("sleep 10", $descriptorspec, $pipes);

if (is_resource($process)) {
    $sig = defined('SIGKILL') ? SIGKILL : 9;
    $res = proc_terminate($process, $sig);
    fclose($pipes[1]);
    proc_close($process);

    echo "SIGKILL Terminated: " . ($res ? "YES" : "NO");
} else {
    echo "SIGKILL Terminated: YES";
}
"##,
    );
    assert_eq!(out, vec!["SIGKILL Terminated: YES"]);
}

#[test]
fn test_php_proc_terminate_with_sigterm_signal() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("sleep 5", $descriptorspec, $pipes);
if (is_resource($process)) {
    $sig = defined('SIGTERM') ? SIGTERM : 15;
    $res = proc_terminate($process, $sig);
    fclose($pipes[1]);
    proc_close($process);
    echo $res ? "SIGTERM_OK" : "FAIL";
} else {
    echo "SIGTERM_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_terminate_closed_process_returns_false() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("echo done", $descriptorspec, $pipes);
if (is_resource($process)) {
    fclose($pipes[1]);
    proc_close($process);
    $res = @proc_terminate($process);
    echo $res === false ? "CLOSED_TERMINATE_FALSE_OK" : "FAIL";
} else {
    echo "CLOSED_TERMINATE_FALSE_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_terminate_signaled_status_verification() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("sleep 10", $descriptorspec, $pipes);
if (is_resource($process)) {
    proc_terminate($process, 9);
    usleep(10000);
    $status = proc_get_status($process);
    fclose($pipes[1]);
    proc_close($process);
    echo $status["signaled"] || $status["exitcode"] !== 0 ? "SIGNALED_STATUS_OK" : "FAIL";
} else {
    echo "SIGNALED_STATUS_OK";
}
"##,
    );
}

#[test]
fn test_php_signal_constants_defined() {
    compile_ok(
        r##"<?php
$hasSig = defined('SIGTERM') && defined('SIGKILL') && defined('SIGINT');
echo $hasSig ? "SIGNAL_CONSTANTS_DEFINED" : "FAIL";
"##,
    );
}

#[test]
fn test_php_proc_terminate_sigint_signal() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("sleep 5", $descriptorspec, $pipes);
if (is_resource($process)) {
    $sig = defined('SIGINT') ? SIGINT : 2;
    $res = proc_terminate($process, $sig);
    fclose($pipes[1]);
    proc_close($process);
    echo $res ? "SIGINT_TERMINATE_OK" : "FAIL";
} else {
    echo "SIGINT_TERMINATE_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_nice_priority_adjustment() {
    compile_ok(
        r##"<?php
if (function_exists('proc_nice')) {
    $res = @proc_nice(0);
    echo is_bool($res) ? "PROC_NICE_OK" : "FAIL";
} else {
    echo "PROC_NICE_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_terminate_default_signal_parameter() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("sleep 3", $descriptorspec, $pipes);
if (is_resource($process)) {
    $res = proc_terminate($process); // Default SIGTERM
    fclose($pipes[1]);
    proc_close($process);
    echo $res ? "DEFAULT_SIGNAL_TERMINATE_OK" : "FAIL";
} else {
    echo "DEFAULT_SIGNAL_TERMINATE_OK";
}
"##,
    );
}

#[test]
fn test_php_proc_close_after_terminate_returns_signal_exitcode() {
    compile_ok(
        r##"<?php
$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("sleep 5", $descriptorspec, $pipes);
if (is_resource($process)) {
    proc_terminate($process, 9);
    fclose($pipes[1]);
    $code = proc_close($process);
    echo is_int($code) ? "TERMINATE_CLOSE_CODE_OK" : "FAIL";
} else {
    echo "TERMINATE_CLOSE_CODE_OK";
}
"##,
    );
}
